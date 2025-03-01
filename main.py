import os
import sys
import glob
import zipfile
import io
from lxml import etree
from openpyxl import Workbook, load_workbook
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from threading import Thread
import sv_ttk  # Custom modern theme for ttk (Sun Valley theme)
import time  # Needed for backup filenames, etc.

# PyInstaller path fix for theme files
if getattr(sys, 'frozen', False):
    sv_ttk_path = os.path.join(sys._MEIPASS, 'sv_ttk')
    sys.path.insert(0, sv_ttk_path)

# -------------------------------------------------------------------------
# 1) Helper functions for collecting XML from ZIPs (including nested zips)
# -------------------------------------------------------------------------

def collect_xml_files_from_zip(zip_path, current_folder_name=None):
    """
    Recursively collect XML files from a ZIP on disk (including nested ZIPs).
    Returns a list of (xml_bytes, folder_name, filename).
    """
    xml_files = []
    
    if current_folder_name is None:
        # Use the ZIP’s base name (without .zip) as the top-level folder name
        base = os.path.basename(zip_path)
        current_folder_name = os.path.splitext(base)[0]
    
    with zipfile.ZipFile(zip_path, 'r') as zf:
        for name in zf.namelist():
            if name.endswith('/'):
                # It's a directory inside the ZIP
                continue
            
            file_bytes = zf.read(name)
            
            if name.lower().endswith('.zip'):
                # Nested ZIP: open in-memory, recurse
                nested_zip = io.BytesIO(file_bytes)
                sub_folder_name = f"{current_folder_name}/{os.path.splitext(os.path.basename(name))[0]}"
                xml_files.extend(collect_xml_files_from_inmemory(nested_zip, sub_folder_name))
            elif name.lower().endswith('.xml'):
                # It's a normal XML
                filename = os.path.basename(name)
                xml_files.append((file_bytes, current_folder_name, filename))
            else:
                # Something else
                pass
    
    return xml_files

def collect_xml_files_from_inmemory(zip_bytes, current_folder_name):
    """
    Same logic as collect_xml_files_from_zip, but the first argument
    is an in-memory BytesIO for a nested ZIP.
    """
    xml_files = []
    with zipfile.ZipFile(zip_bytes, 'r') as zf:
        for name in zf.namelist():
            if name.endswith('/'):
                continue
            
            file_bytes = zf.read(name)
            if name.lower().endswith('.zip'):
                # Another nested ZIP
                nested_zip = io.BytesIO(file_bytes)
                sub_folder_name = f"{current_folder_name}/{os.path.splitext(os.path.basename(name))[0]}"
                xml_files.extend(collect_xml_files_from_inmemory(nested_zip, sub_folder_name))
            elif name.lower().endswith('.xml'):
                filename = os.path.basename(name)
                xml_files.append((file_bytes, current_folder_name, filename))
    return xml_files


# -------------------------------------------------------------------------
# 2) Main Application
# -------------------------------------------------------------------------
class ModernInvoiceProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("XML Invoice Processor")
        self.root.geometry("700x500")
        self.root.minsize(600, 400)
        
        # Apply Sun Valley theme (modern Windows 11 style)
        sv_ttk.set_theme("dark")  # Can be "light" or "dark"
        
        # Configure the root grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # Create main frame with padding
        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.main_frame.columnconfigure(0, weight=1)
        
        # App header with logo/icon
        self.header_frame = ttk.Frame(self.main_frame)
        self.header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 25))
        self.header_frame.columnconfigure(1, weight=1)
        
        # Create a simple icon/logo
        self.canvas = tk.Canvas(self.header_frame, width=40, height=40, 
                               highlightthickness=0, bg=self.root.cget('bg'))
        self.canvas.grid(row=0, column=0, padx=(0, 15))
        
        # Draw a simple document icon
        self.canvas.create_rectangle(10, 5, 30, 35, fill="#4F9BF8", outline="")
        self.canvas.create_rectangle(15, 10, 25, 15, fill="white", outline="")
        self.canvas.create_rectangle(15, 18, 25, 23, fill="white", outline="")
        self.canvas.create_rectangle(15, 26, 25, 31, fill="white", outline="")
        
        # Header text
        self.header_label = ttk.Label(self.header_frame, 
                                     text="XML Invoice Processor", 
                                     font=("Segoe UI", 18, "bold"))
        self.header_label.grid(row=0, column=1, sticky="w")
        
        # Card-like container for inputs
        self.card_frame = ttk.LabelFrame(self.main_frame, text="Configuration")
        self.card_frame.grid(row=1, column=0, sticky="ew", pady=(0, 20), ipady=15)
        self.card_frame.columnconfigure(0, weight=1)
        
        # Folder selection section with improved layout
        self.folder_frame = ttk.Frame(self.card_frame, padding="10")
        self.folder_frame.grid(row=0, column=0, sticky="ew", padx=10)
        self.folder_frame.columnconfigure(1, weight=1)
        
        # Folder label
        self.folder_label = ttk.Label(self.folder_frame, 
                                     text="XML Invoices Folder or ZIP:", 
                                     font=("Segoe UI", 10))
        self.folder_label.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
        
        # Input container for folder
        self.folder_input_frame = ttk.Frame(self.folder_frame)
        self.folder_input_frame.grid(row=1, column=0, columnspan=2, sticky="ew")
        self.folder_input_frame.columnconfigure(0, weight=1)
        
        # Folder path entry
        self.folder_path = tk.StringVar()
        self.folder_entry = ttk.Entry(self.folder_input_frame, textvariable=self.folder_path)
        self.folder_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        
        # Browse button
        self.browse_button = ttk.Button(self.folder_input_frame, text="Browse", 
                                       command=self.browse_folder, width=15)
        self.browse_button.grid(row=0, column=1, sticky="e")
        
        # Separator
        self.separator = ttk.Separator(self.card_frame, orient="horizontal")
        self.separator.grid(row=1, column=0, sticky="ew", pady=15, padx=20)
        
        # Output file section
        self.output_frame = ttk.Frame(self.card_frame, padding="10")
        self.output_frame.grid(row=2, column=0, sticky="ew", padx=10)
        self.output_frame.columnconfigure(1, weight=1)
        
        # Output file label
        self.output_label = ttk.Label(self.output_frame, 
                                     text="Output Excel File:", 
                                     font=("Segoe UI", 10))
        self.output_label.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
        
        # Input container for output
        self.output_input_frame = ttk.Frame(self.output_frame)
        self.output_input_frame.grid(row=1, column=0, columnspan=2, sticky="ew")
        self.output_input_frame.columnconfigure(0, weight=1)
        
        # Output file entry
        self.output_path = tk.StringVar(value="All_Invoices.xlsx")
        self.output_entry = ttk.Entry(self.output_input_frame, textvariable=self.output_path)
        self.output_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        
        # Output browse button
        self.output_button = ttk.Button(self.output_input_frame, text="Browse", 
                                       command=self.browse_output, width=15)
        self.output_button.grid(row=0, column=1, sticky="e")
        
        # Action section frame
        self.action_frame = ttk.Frame(self.main_frame)
        self.action_frame.grid(row=2, column=0, sticky="ew", pady=(0, 20))
        self.action_frame.columnconfigure(0, weight=1)
        
        # Process button
        self.process_button = ttk.Button(self.action_frame, 
                                        text="Process Invoices", 
                                        command=self.process_invoices, 
                                        style="Accent.TButton",
                                        width=25)
        self.process_button.grid(row=0, column=0, pady=10)
        
        # Status section
        self.status_frame = ttk.LabelFrame(self.main_frame, text="Status")
        self.status_frame.grid(row=3, column=0, sticky="ew", pady=(0, 10), ipady=5)
        self.status_frame.columnconfigure(0, weight=1)
        
        # Progress frame
        self.progress_frame = ttk.Frame(self.status_frame, padding="10")
        self.progress_frame.grid(row=0, column=0, sticky="ew")
        self.progress_frame.columnconfigure(0, weight=1)
        
        # Progress bar
        self.progress = ttk.Progressbar(self.progress_frame, mode='indeterminate')
        self.progress.grid(row=0, column=0, sticky="ew", pady=(5, 10))
        
        # Status label
        self.status_var = tk.StringVar(value="Ready to process invoices")
        self.status_label = ttk.Label(self.progress_frame, 
                                     textvariable=self.status_var, 
                                     font=("Segoe UI", 9))
        self.status_label.grid(row=1, column=0, sticky="w", pady=(0, 5))
        
        # Namespaces for XML parsing
        self.ns = {
            "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
            "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"
        }
        
        # Processing flag
        self.processing = False
        
        # Logs section
        self.log_frame = ttk.LabelFrame(self.main_frame, text="Processing Logs")
        self.log_frame.grid(row=4, column=0, sticky="nsew", pady=(10, 0))
        self.log_frame.columnconfigure(0, weight=1)
        self.log_frame.rowconfigure(0, weight=1)
        
        # Log text widget
        self.log_text = tk.Text(self.log_frame, height=5, wrap="word", 
                               font=("Consolas", 9))
        self.log_text.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        # Scrollbar for logs
        self.log_scrollbar = ttk.Scrollbar(self.log_frame, orient="vertical", 
                                          command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=self.log_scrollbar.set)
        self.log_scrollbar.grid(row=0, column=1, sticky="ns", padx=(0, 10), pady=10)
        
        # Configure main frame rows to expand log section
        self.main_frame.rowconfigure(4, weight=1)

    # ---------------------------------------------------------------------
    # GUI Helpers
    # ---------------------------------------------------------------------
    
    def browse_folder(self):
        """
        Single 'Browse' that tries to pick either a folder or a ZIP file.
        On many systems, askopenfilename won't let you pick a directory easily.
        If you prefer old behavior, revert to askdirectory for folders.
        """
        path = filedialog.askopenfilename(
            title="Select a folder or a ZIP containing XML invoices",
            filetypes=[("All files", "*.*")]
        )
        
        if not path:
            return  # User canceled
        
        # If path is actually a directory, or a .zip, accept it. Otherwise, show error.
        if os.path.isdir(path):
            self.folder_path.set(path)
        elif path.lower().endswith('.zip'):
            self.folder_path.set(path)
        else:
            messagebox.showerror("Invalid selection", "Please select a folder OR a .zip file.")
    
    def browse_output(self):
        output_path = filedialog.asksaveasfilename(
            title="Save Excel File",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if output_path:
            self.output_path.set(output_path)
    
    def log(self, message):
        """Add message to log with timestamp"""
        import datetime
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.root.after(0, lambda: self.log_text.insert(tk.END, f"[{timestamp}] {message}\n"))
        self.root.after(0, lambda: self.log_text.see(tk.END))
    
    # ---------------------------------------------------------------------
    # Main Processing
    # ---------------------------------------------------------------------
    
    def process_invoices(self):
        if self.processing:
            return
        
        folder_path = self.folder_path.get().strip()
        output_path = self.output_path.get().strip()
        
        if not folder_path:
            messagebox.showerror("Error", "Please select a folder or .zip file containing invoices")
            return
        
        if not output_path:
            messagebox.showerror("Error", "Please specify an output Excel file")
            return
        
        # Clear log
        self.log_text.delete(1.0, tk.END)
        
        # Start processing in a thread
        self.processing = True
        self.progress.start()
        self.status_var.set("Processing invoices... Please wait")
        self.process_button.state(['disabled'])
        
        self.log(f"Starting to process invoices from: {folder_path}")
        self.log(f"Output will be saved to: {output_path}")
        
        process_thread = Thread(target=self.process_invoices_thread, args=(folder_path, output_path))
        process_thread.daemon = True
        process_thread.start()
    
    def get_existing_invoices(self, output_path):
        """Load existing invoices from Excel file to avoid duplicates"""
        existing_invoices = set()
        try:
            if os.path.exists(output_path):
                wb = load_workbook(output_path)
                ws = wb.active
                
                # Get the column index for invoice number
                invoice_col_idx = None
                for idx, cell in enumerate(ws[1]):
                    if cell.value == "Nr. doc.":
                        invoice_col_idx = idx
                        break
                
                if invoice_col_idx is not None:
                    # Start from row 2 (skipping headers)
                    for row in list(ws.rows)[1:]:
                        invoice_num = row[invoice_col_idx].value
                        if invoice_num:
                            existing_invoices.add(invoice_num)
                
                self.log(f"Found {len(existing_invoices)} existing invoices in the Excel file")
                return existing_invoices, wb, ws
            else:
                self.log("No existing Excel file found, will create a new one")
                return set(), None, None
        except Exception as e:
            self.log(f"Error reading existing Excel file: {e}")
            return set(), None, None
    
    def process_invoices_thread(self, folder_path, output_path):
        try:
            # Check if the output file exists and get existing invoice numbers
            existing_invoices, existing_wb, existing_ws = self.get_existing_invoices(output_path)
            
            # If we don't have an existing workbook, create a new one
            if existing_wb is None:
                wb = Workbook()
                ws = wb.active
                
                # Define the column headers
                ws.append([
                    "Folder Name",           # Folder containing the invoice
                    "Nr. doc.",              # cbc:ID
                    "Data emiterii",         # cbc:IssueDate
                    "Termen plata",          # cbc:DueDate
                    "Cota TVA",              # from <cac:TaxSubtotal><cac:TaxCategory><cbc:Percent>
                    "Furnizor",              # Supplier name
                    "CIF",                   # Supplier tax ID
                    "Reg. com.",             # Supplier trade register ID
                    "Adresa",                # Supplier street
                    "Judet",                 # Supplier region
                    "IBAN",                  # PaymentMeans -> PayeeFinancialAccount -> cbc:ID
                    "Banca",                 # PaymentMeans -> PayeeFinancialAccount -> cbc:Name
                    "Produse/Servicii",      # <cac:Item><cbc:Name>
                    "Descriere",             # <cac:Item><cbc:Description>
                    "U.M.",                  # <cbc:InvoicedQuantity unitCode="??">
                    "Cant.",                 # <cbc:InvoicedQuantity> text
                    "Pret fara TVA (RON)",   # <cac:Price><cbc:PriceAmount>
                    "Valoare",               # <cac:InvoiceLine><cbc:LineExtensionAmount> (net)
                    "Valoare TVA",           # line net * cota TVA
                    "Total",                 # net + tax for the line
                    "Total factura"          # <cac:LegalMonetaryTotal><cbc:PayableAmount> (invoice total)
                ])
            else:
                wb = existing_wb
                ws = existing_ws
            
            file_count = 0
            error_count = 0
            skipped_count = 0
            new_invoices_count = 0
            
            # -------------------------------------------------------------
            # A) Gather all XML files (in-memory) from the user selection
            # -------------------------------------------------------------
            xml_files = []  # will hold (xml_content, folder_name, filename)
            
            if os.path.isdir(folder_path):
                # Normal folder logic
                for root_dir, dirs, files in os.walk(folder_path):
                    folder_name = os.path.basename(root_dir)
                    
                    for filename in files:
                        # Skip non-XML or signature files
                        if not filename.lower().endswith(".xml"):
                            continue
                        if "semnatura" in filename.lower():
                            continue
                        
                        file_path = os.path.join(root_dir, filename)
                        with open(file_path, "rb") as f:
                            xml_content = f.read()
                        xml_files.append((xml_content, folder_name, filename))
            
            else:
                # It's presumably a ZIP file
                if folder_path.lower().endswith(".zip"):
                    self.log(f"Detected ZIP file: {folder_path}")
                    xml_files_in_zip = collect_xml_files_from_zip(folder_path)
                    # each item is (xml_bytes, folder_name, filename)
                    xml_files.extend(xml_files_in_zip)
                else:
                    # Not a dir, not a zip -> can't process
                    self.log("ERROR: The selected path is not a folder and not a .zip")
                    self.root.after(0, lambda: self.finish_processing(0, 0, 0, False))
                    return
            
            total_files = len(xml_files)
            self.log(f"Found {total_files} XML files to process")
            
            # -------------------------------------------------------------
            # B) Parse each XML file (now in memory) and write to Excel
            # -------------------------------------------------------------
            
            for (xml_content, folder_name, filename) in xml_files:
                try:
                    # Use lxml from an in-memory BytesIO
                    tree = etree.parse(io.BytesIO(xml_content))
                    root_xml = tree.getroot()
                    
                    # Get invoice number to check for duplicates
                    invoice_number = self.get_text(root_xml.find(".//cbc:ID", self.ns))
                    
                    if invoice_number in existing_invoices:
                        skipped_count += 1
                        self.log(f"Skipping existing invoice: {invoice_number}")
                        continue
                    
                    issue_date = self.get_text(root_xml.find(".//cbc:IssueDate", self.ns))
                    due_date = self.get_text(root_xml.find(".//cbc:DueDate", self.ns))
                    
                    # Supplier info
                    supplier_name = self.get_text(root_xml.find(
                        ".//cac:AccountingSupplierParty//cac:PartyLegalEntity/cbc:RegistrationName", self.ns))
                    supplier_cif = self.get_text(root_xml.find(
                        ".//cac:AccountingSupplierParty//cac:PartyTaxScheme/cbc:CompanyID", self.ns))
                    supplier_regcom = self.get_text(root_xml.find(
                        ".//cac:AccountingSupplierParty//cac:PartyLegalEntity/cbc:CompanyID", self.ns))
                    supplier_street = self.get_text(root_xml.find(
                        ".//cac:AccountingSupplierParty//cac:PostalAddress/cbc:StreetName", self.ns))
                    supplier_judet = self.get_text(root_xml.find(
                        ".//cac:AccountingSupplierParty//cac:PostalAddress/cbc:CountrySubentity", self.ns))
                    
                    # Invoice-level Cota TVA
                    tax_percent_el = root_xml.find(".//cac:TaxSubtotal/cac:TaxCategory/cbc:Percent", self.ns)
                    invoice_vat_percent_text = self.get_text(tax_percent_el)
                    
                    # Payment details
                    payee_account = root_xml.find(".//cac:PaymentMeans/cac:PayeeFinancialAccount", self.ns)
                    iban = self.get_text(payee_account.find("cbc:ID", self.ns)) if payee_account is not None else ""
                    bank_name = self.get_text(payee_account.find("cbc:Name", self.ns)) if payee_account is not None else ""
                    
                    # Total factura
                    total_invoice = self.get_text(root_xml.find(".//cac:LegalMonetaryTotal/cbc:PayableAmount", self.ns))
                    
                    # Invoice lines
                    invoice_lines = root_xml.findall(".//cac:InvoiceLine", self.ns)
                    if not invoice_lines:
                        # No lines -> append a blank line row
                        ws.append([
                            folder_name,
                            invoice_number,
                            issue_date,
                            due_date,
                            invoice_vat_percent_text,
                            supplier_name,
                            supplier_cif,
                            supplier_regcom,
                            supplier_street,
                            supplier_judet,
                            iban,
                            bank_name,
                            "",  # Produse/Servicii
                            "",  # Descriere
                            "",  # U.M.
                            "",  # Cant.
                            "",  # Pret fara TVA
                            "",  # Valoare
                            "",  # Valoare TVA
                            "",  # Total
                            total_invoice
                        ])
                        new_invoices_count += 1
                    else:
                        # Otherwise, one row per line
                        for line in invoice_lines:
                            # Quantity & unit
                            invoiced_qty = line.find("./cbc:InvoicedQuantity", self.ns)
                            quantity_text = self.get_text(invoiced_qty)
                            uom = invoiced_qty.get("unitCode") if invoiced_qty is not None else ""
                            
                            # Unit price
                            price_el = line.find("./cac:Price/cbc:PriceAmount", self.ns)
                            price_text = self.get_text(price_el)
                            
                            # Net line extension
                            line_ext_el = line.find("./cbc:LineExtensionAmount", self.ns)
                            line_ext_text = self.get_text(line_ext_el)
                            
                            # Calculate line tax & total using invoice-level VAT
                            try:
                                line_ext = float(line_ext_text) if line_ext_text else 0.0
                                vat_percent = float(invoice_vat_percent_text) if invoice_vat_percent_text else 0.0
                                line_tax = round(line_ext * vat_percent / 100, 2)
                                line_total = round(line_ext + line_tax, 2)
                            except ValueError:
                                line_tax = ""
                                line_total = ""
                            
                            # Product/Service name & description
                            item_name = self.get_text(line.find("./cac:Item/cbc:Name", self.ns))
                            item_desc = self.get_text(line.find("./cac:Item/cbc:Description", self.ns))
                            
                            # Append row
                            ws.append([
                                folder_name,
                                invoice_number,
                                issue_date,
                                due_date,
                                invoice_vat_percent_text,
                                supplier_name,
                                supplier_cif,
                                supplier_regcom,
                                supplier_street,
                                supplier_judet,
                                iban,
                                bank_name,
                                item_name,        
                                item_desc,        
                                uom,              
                                quantity_text,    
                                price_text,       
                                line_ext_text,    
                                line_tax,         
                                line_total,       
                                total_invoice     
                            ])
                        
                        new_invoices_count += 1
                    
                    existing_invoices.add(invoice_number)
                    file_count += 1
                    
                    # Update status in main thread
                    progress_msg = f"Processed {file_count} files, skipped {skipped_count} duplicates..."
                    self.root.after(0, lambda msg=progress_msg: self.status_var.set(msg))
                    self.log(f"Processed: {filename} - Invoice: {invoice_number}")
                    
                except Exception as e:
                    error_count += 1
                    self.log(f"Error parsing {filename}: {str(e)}")
            
            # Check if we found any new invoices
            if new_invoices_count == 0 and skipped_count > 0:
                self.log("No new invoices found, Excel file remains unchanged")
                self.root.after(0, lambda: self.finish_processing(file_count, error_count, skipped_count, False))
                return
            
            # Apply column widths for better readability (first row is header)
            for col_idx, column in enumerate(ws.columns, 1):
                max_length = 0
                column_name = column[0].value
                # Check first 100 rows for performance
                for i, cell in enumerate(column[:100]):
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                
                adjusted_width = max(max_length, len(str(column_name))) + 3
                ws.column_dimensions[chr(64 + col_idx)].width = min(adjusted_width, 35)  # Cap at width 35
            
            # Create a backup if file exists
            try:
                if os.path.exists(output_path):
                    backup_path = f"{os.path.splitext(output_path)[0]}_backup_{int(time.time())}.xlsx"
                    import shutil
                    shutil.copy2(output_path, backup_path)
                    self.log(f"Created backup at: {os.path.basename(backup_path)}")
            except Exception as e:
                self.log(f"Warning: Could not create backup: {str(e)}")
            
            # Save the Excel file
            try:
                wb.save(output_path)
                self.log(f"Successfully saved Excel file: {output_path}")
            except PermissionError:
                self.log("ERROR: Could not save Excel file - it might be open in another program")
                self.root.after(0, lambda: self.show_error("Could not save Excel file. Please close it if it's open in another program."))
                return
            except Exception as e:
                self.log(f"ERROR: Failed to save Excel file: {str(e)}")
                self.root.after(0, lambda: self.show_error(f"Failed to save Excel file: {str(e)}"))
                return
            
            self.root.after(0, lambda: self.finish_processing(file_count, error_count, skipped_count, True))
        
        except Exception as e:
            import traceback
            self.log(f"CRITICAL ERROR: {str(e)}")
            self.log(traceback.format_exc())
            self.root.after(0, lambda: self.show_error(str(e)))
    
    def get_text(self, element):
        """Safely return element.text or empty string."""
        return element.text if element is not None else ""
    
    def finish_processing(self, file_count, error_count, skipped_count, saved):
        self.progress.stop()
        self.processing = False
        self.process_button.state(['!disabled'])
        
        if not saved:
            message = f"No new invoices found. Processed {file_count} files, all {skipped_count} were duplicates."
            self.status_var.set(message)
            messagebox.showinfo("Processing Complete", message)
            return
        
        if error_count > 0:
            message = f"Done! Processed {file_count} files ({skipped_count} duplicates) with {error_count} errors."
        else:
            message = f"Done! Processed {file_count} files successfully, skipped {skipped_count} duplicates."
        
        self.status_var.set(message)
        messagebox.showinfo("Processing Complete", f"{message}\nExcel file saved to:\n{self.output_path.get()}")
    
    def show_error(self, error_message):
        self.progress.stop()
        self.processing = False
        self.process_button.state(['!disabled'])
        self.status_var.set("Error during processing")
        messagebox.showerror("Error", f"An error occurred:\n{error_message}")

# -------------------------------------------------------------------------
# Main entry point
# -------------------------------------------------------------------------
if __name__ == "__main__":
    root = tk.Tk()
    # If sv_ttk is missing, try to install it once
    try:
        import sv_ttk
    except ImportError:
        import subprocess
        messagebox.showinfo("First Run Setup", 
                           "Installing the modern theme component (sv_ttk).\nThis will only happen once.")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "sv-ttk"])
        import sv_ttk
    else:
        pass
    
    app = ModernInvoiceProcessorApp(root)
    root.mainloop()
