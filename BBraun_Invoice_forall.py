import pdfplumber
import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk

class InvoiceProcessor:
    def __init__(self):
        # Patterns for Format 1 & 2
        self.qty_p1 = re.compile(r"([\d,.]+\d{2}|[\d,.]+)\s+(PC|PAC|PCS)\b")
        self.origin_p1 = re.compile(r"\bN\s+([A-Z]{2})\b")
        
        # Patterns for Format 3
        self.qty_p3 = re.compile(r"([\d,]+\.\d{2})\s+(PC|PAC|PCS)")
        self.price_p3 = re.compile(r"([\d,]+\.\d{2})\s*(?=/|EUR|1\s*PC|1\s*PAC)")

    def extract_format_1_2(self, pages):
        """Logic for in_mumbai & Druckdaten (004) GERMANY SPAIN"""
        all_rows = []
        current_item = None
        inv_no, inv_date = None, None

        for page in pages:
            text = page.extract_text()
            if not text: continue
            lines = text.split('\n')

            for i, line in enumerate(lines):
                line = line.strip()
                parts = line.split()
                if not parts: continue

                # Header Logic
                clean_line = re.sub(r'[^a-zA-Z0-9.\s:]', '', line)
                if not inv_no and any(x in clean_line.lower() for x in ["document", "doc", "invoice"]):
                    num_match = re.search(r'(\d{8,10})', clean_line)
                    if num_match: inv_no = num_match.group(1)

                if not inv_date:
                    date_match = re.search(r"\d{2}\.\d{2}\.\d{4}", clean_line)
                    if date_match: inv_date = date_match.group()

                # Item Trigger (Checks if first part is a position number like 0010)
                if parts[0].isdigit() and len(parts) > 1:
                    pos_len = len(parts[0])
                    if pos_len in [1, 2, 3, 4, 6]: # Expanded range
                        potential_code = parts[1]

                        # Ignore common address keywords
                        if not any(city in line.upper() for city in ["MUMBAI", "BHIWANDI", "MELSUNGEN", "GERMANY"]):
                            if current_item: 
                                all_rows.append(current_item)
                            
                            current_item = {
                                "Invoice Number": inv_no, 
                                "Inv Date": inv_date, 
                                "Code": potential_code,
                                "Description": "", 
                                "Origin Code": "", 
                                "Quantity": "", 
                                "Unit": "", 
                                "Price": ""
                            }

                            # Same-line extraction
                            orig_m = self.origin_p1.search(line)
                            if orig_m: current_item["Origin Code"] = orig_m.group(1)
                            
                            q_m = self.qty_p1.search(line)
                            if q_m:
                                raw_qty = q_m.group(1).strip()
                                current_item["Quantity"] = raw_qty
                                current_item["Unit"] = q_m.group(2).strip()
                                
                                # Find Price after the quantity
                                line_after_qty = line.split(q_m.group())[-1].strip()
                                p_match = re.search(r"([\d,.]+)(?=\s*(?:\d\s+)?(?:PC|PAC|EUR|PCS|\b))", line_after_qty)
                                if p_match: current_item["Price"] = p_match.group(1)

                            # Description Extraction
                            desc_text = " ".join(parts[2:])
                            clean_desc = re.split(r'\bN\s+[A-Z]{2}\b|\d+\s+(PC|PAC|PCS)', desc_text)[0]
                            current_item["Description"] = clean_desc.strip()
                            
                            # Fallback to next line if description is empty
                            if current_item["Description"] == "" and i + 1 < len(lines):
                                next_line = lines[i+1].strip()
                                if not next_line[0:2].isdigit(): # Ensure it's not another item
                                    current_item["Description"] = next_line

                # Separate line Price check (only if price wasn't found on same line)
                if current_item and not current_item["Price"]:
                    if "Price w/o Dto" in line or "EUR" in line:
                        p_match = re.search(r"([\d,.]+)\s*EUR", line)
                        if p_match: current_item["Price"] = p_match.group(1)

        if current_item: 
            all_rows.append(current_item)
        return all_rows

    def extract_format_3(self, pages):
        """Logic for inv-872264189 MALAYSIA"""
        all_rows = []
        current_item = None
        inv_no, inv_date = None, None

        for page in pages:
            lines = page.extract_text().split('\n')
            for i, line in enumerate(lines):
                line = line.strip()
                parts = line.split()
                if not parts: continue

                clean_line = re.sub(r'[^a-zA-Z0-9.\s:]', '', line)
                if not inv_no:
                    h_match = re.search(r'(?:Document|Doc)\.?\s*(?:no\.?)?\s*(\d{8,})', clean_line, re.IGNORECASE)
                    if h_match: inv_no = h_match.group(1)
                if not inv_date:
                    d_match = re.search(r"(\d{4}-\d{2}-\d{2}|\d{2}\.\d{2}\.\d{4})", line)
                    if d_match: inv_date = d_match.group()

                if parts[0].isdigit() and (len(parts[0]) in [2, 3, 4, 6]):
                    potential_code = parts[1] if len(parts) > 1 else ""
                    if any(city in potential_code.upper() for city in ["MUMBAI", "BHIWANDI", "PENANG"]): continue
                    
                    if current_item: all_rows.append(current_item)
                    desc_full = " ".join(parts[2:]) if len(parts) > 2 else ""
                    desc_clean = re.split(r'\b[\d,.]+\s+(?:PC|PAC|PCS)\b|HS Code|DO No', desc_full)[0].strip()
                    current_item = {
                        "Invoice Number": inv_no, "Inv Date": inv_date, "Code": potential_code,
                        "Description": desc_clean,
                        "Origin Code": "", "Quantity": "", "Price": ""
                    }
                    q_m = self.qty_p3.search(line)
                    if q_m: 
                        current_item["Quantity"] = q_m.group(1).replace(',', '')
                        current_item["Unit"] = q_m.group(2)
                    p_m = self.price_p3.search(line)
                    if p_m: current_item["Price"] = p_m.group(1).replace(',', '')
                    continue

                if current_item:
                    if any(k in line for k in ["Insurance", "Freight", "Total", "SDN. BHD."]):
                        all_rows.append(current_item)
                        current_item = None
                        continue
                    if not current_item["Quantity"]:
                        q_m = self.qty_p3.search(line)
                        if q_m: 
                            current_item["Quantity"] = q_m.group(1).replace(',', '')
                            current_item["Unit"] = q_m.group(2)
                    if not current_item["Price"]:
                        p_m = self.price_p3.search(line)
                        if p_m: current_item["Price"] = p_m.group(1).replace(',', '')

        if current_item: all_rows.append(current_item)
        return all_rows

# --- GUI APPLICATION ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Nagarkot Forwarders - Invoice Processor")
        self.root.geometry("900x700")
        self.root.configure(bg="#ffffff") # White background

        self.processor = InvoiceProcessor()
        self.file_paths = []
        self.format_choice = tk.IntVar(value=1)

        # --- THEME SETUP ---
        self.setup_styles()

        # --- LAYOUT ---
        # 1. Header
        self.create_header()

        # 2. Main Content
        self.create_main_content()

        # 3. Footer
        self.create_footer()

    def setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        # General Colors
        bg_color = "#ffffff"
        fg_color = "#333333"
        brand_blue = "#0056b3"
        
        self.style.configure("TFrame", background=bg_color)
        self.style.configure("TLabel", background=bg_color, foreground=fg_color, font=("Arial", 10))
        self.style.configure("TButton", font=("Arial", 10, "bold"))
        self.style.configure("TLabelframe", background=bg_color)
        self.style.configure("TLabelframe.Label", background=bg_color, foreground=brand_blue, font=("Arial", 10, "bold"))
        
        # Header Styles
        self.style.configure("Header.TLabel", font=("Helvetica", 18, "bold"), foreground=brand_blue)
        self.style.configure("SubHeader.TLabel", font=("Helvetica", 10), foreground="#888888")

        # Buttons
        # Primary (Blue)
        self.style.configure("Primary.TButton", background=brand_blue, foreground="white", borderwidth=0)
        self.style.map("Primary.TButton", background=[("active", "#004494")])
        
        # Secondary (White/Gray)
        self.style.configure("Secondary.TButton", background="#f0f0f0", foreground=fg_color, bordercolor="#cccccc")
        self.style.map("Secondary.TButton", background=[("active", "#e0e0e0")])

    def create_header(self):
        # 3-Column Layout: Logo (Left), Title (Center), Spacer (Right)
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill="x", padx=20, pady=15)
        header_frame.columnconfigure(1, weight=1) # Center expands

        # LOGO
        try:
            # Load from current root directory
            img = Image.open("Nagarkot Logo.png")
            # Resize to Height = 20px
            h_percent = (20 / float(img.size[1]))
            w_size = int((float(img.size[0]) * float(h_percent)))
            img = img.resize((w_size, 20), Image.Resampling.LANCZOS)
            
            self.logo_img = ImageTk.PhotoImage(img) # Keep reference
            logo_label = ttk.Label(header_frame, image=self.logo_img)
            logo_label.grid(row=0, column=0, sticky="w")
        except Exception as e:
            print(f"Warning: Nagarkot Logo.png not found or invalid ({e})")
            # Text Placeholder
            ttk.Label(header_frame, text="[NAGARKOT LOGO]", font=("Arial", 10, "bold"), foreground="gray").grid(row=0, column=0, sticky="w")

        # TITLE (Centered)
        title_container = ttk.Frame(header_frame)
        title_container.grid(row=0, column=1) # Center column
        
        title_label = ttk.Label(title_container, text="INVOICE PROCESSOR", style="Header.TLabel")
        title_label.pack(anchor="center")
        
        subtitle_label = ttk.Label(title_container, text="Automated Extraction Tool", style="SubHeader.TLabel")
        subtitle_label.pack(anchor="center")

        # Right spacer (optional)
        ttk.Label(header_frame, text="   ").grid(row=0, column=2, sticky="e")

    def create_main_content(self):
        main_content = ttk.Frame(self.root, padding="20")
        main_content.pack(fill="both", expand=True)

        # --- INPUT SECTION (LabelFrame) ---
        input_frame = ttk.LabelFrame(main_content, text="Selection & Options", padding="15")
        input_frame.pack(fill="x", pady=(0, 15))

        # File Selection
        ttk.Label(input_frame, text="Selected PDF Files:").pack(anchor="w", pady=(0, 5))
        
        file_toolbar = ttk.Frame(input_frame)
        file_toolbar.pack(fill="x", pady=(0, 5))
        
        ttk.Button(file_toolbar, text="📂 Select Files", command=self.browse_files, style="Secondary.TButton").pack(side="left", padx=(0, 5))
        ttk.Button(file_toolbar, text="❌ Clear List", command=self.clear_files, style="Secondary.TButton").pack(side="left")

        # Listbox
        list_frame = ttk.Frame(input_frame)
        list_frame.pack(fill="x", pady=5)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        
        self.file_listbox = tk.Listbox(list_frame, height=5, yscrollcommand=scrollbar.set, selectmode=tk.EXTENDED, bd=1, relief="solid")
        self.file_listbox.pack(side="left", fill="x", expand=True)
        scrollbar.config(command=self.file_listbox.yview)

        self.file_count_lbl = ttk.Label(input_frame, text="0 files selected", font=("Arial", 9), foreground="gray")
        self.file_count_lbl.pack(anchor="w")

        # Format Selection
        ttk.Label(input_frame, text="Invoice Format:").pack(anchor="w", pady=(10, 5))
        
        format_frame = ttk.Frame(input_frame)
        format_frame.pack(fill="x")
        ttk.Radiobutton(format_frame, text="Format 1 (Germany / Spain)", variable=self.format_choice, value=1).pack(side="left", padx=(0, 15))
        ttk.Radiobutton(format_frame, text="Format 2 (Malaysia)", variable=self.format_choice, value=2).pack(side="left")

        # --- DATA PREVIEW SECTION ---
        preview_frame = ttk.LabelFrame(main_content, text="Data Preview", padding="10")
        preview_frame.pack(fill="both", expand=True)
        
        # Treeview
        cols = ("Invoice Number", "Date", "Code", "Quantity", "Price", "Source File")
        self.tree = ttk.Treeview(preview_frame, columns=cols, show="headings", selectmode="browse")
        
        # Setup columns
        self.tree.heading("Invoice Number", text="Inv No")
        self.tree.column("Invoice Number", width=100)
        self.tree.heading("Date", text="Date")
        self.tree.column("Date", width=100)
        self.tree.heading("Code", text="Item Code")
        self.tree.column("Code", width=120)
        self.tree.heading("Quantity", text="Qty")
        self.tree.column("Quantity", width=80)
        self.tree.heading("Price", text="Price")
        self.tree.column("Price", width=80)
        self.tree.heading("Source File", text="File")
        self.tree.column("Source File", width=150)
        
        scrollbar_tree = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar_tree.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar_tree.pack(side="right", fill="y")

    def create_footer(self):
        footer_frame = ttk.Frame(self.root, padding="10")
        footer_frame.pack(side="bottom", fill="x")
        
        # Border top effect using a separator
        ttk.Separator(self.root, orient="horizontal").pack(side="bottom", fill="x", pady=(0, 0))

        # Copyright
        ttk.Label(footer_frame, text="© Nagarkot Forwarders Pvt Ltd", font=("Arial", 9), foreground="gray").pack(side="left")
        
        # Run Button
        self.run_btn = ttk.Button(footer_frame, text="Generate Invoice Data", command=self.process_data, style="Primary.TButton")
        self.run_btn.pack(side="right", padx=10)
        
        # Status Label (next to button)
        self.status_var = tk.StringVar(value="Ready")
        self.status_lbl = ttk.Label(footer_frame, textvariable=self.status_var, foreground="#555555")
        self.status_lbl.pack(side="right", padx=10)

    # --- FUNCTIONALITY ---

    def browse_files(self):
        filenames = filedialog.askopenfilenames(
            title="Select PDF Invoices",
            filetypes=[("PDF files", "*.pdf")]
        )
        if filenames:
            for f in filenames:
                if f not in self.file_paths:
                    self.file_paths.append(f)
                    self.file_listbox.insert(tk.END, os.path.basename(f))
            self.update_count()

    def clear_files(self):
        self.file_paths = []
        self.file_listbox.delete(0, tk.END)
        self.tree.delete(*self.tree.get_children()) # Clear preview too
        self.update_count()
    
    def update_count(self):
        self.file_count_lbl.config(text=f"{len(self.file_paths)} file(s) selected")

    def process_data(self):
        if not self.file_paths:
            messagebox.showwarning("Warning", "Please select at least one PDF file.")
            return

        self.status_var.set("Processing...")
        self.root.update_idletasks()
        
        # Clear previous tree data
        self.tree.delete(*self.tree.get_children())
        
        try:
            all_data = []
            processed_count = 0
            failed_files = []

            for pdf_path in self.file_paths:
                try:
                    if not os.path.exists(pdf_path): continue
                    
                    with pdfplumber.open(pdf_path) as pdf:
                        pages = pdf.pages
                        if self.format_choice.get() == 1:
                            raw_data = self.processor.extract_format_1_2(pages)
                        else:
                            raw_data = self.processor.extract_format_3(pages)
                    
                    if raw_data:
                        for row in raw_data:
                            row["Source File"] = os.path.basename(pdf_path)
                            
                            # Insert into Treeview
                            self.tree.insert("", "end", values=(
                                row.get("Invoice Number", ""),
                                row.get("Inv Date", ""),
                                row.get("Code", ""),
                                row.get("Quantity", ""),
                                row.get("Price", ""),
                                row.get("Source File", "")
                            ))
                            
                        all_data.extend(raw_data)
                        processed_count += 1
                        
                except Exception as e:
                    failed_files.append(f"{os.path.basename(pdf_path)}")
                    print(f"Error processing {pdf_path}: {e}")

            if not all_data:
                self.status_var.set("No data found.")
                messagebox.showinfo("Info", "No data extracted from selected files.")
                return

            # SAVE LOGIC
            from datetime import datetime
            output_folder = os.path.dirname(self.file_paths[0])
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            format_name = "Germany_Spain" if self.format_choice.get() == 1 else "Malaysia"
            output_filename = f"Combined_{format_name}_{len(self.file_paths)}files_{timestamp}.xlsx"
            output_path = os.path.join(output_folder, output_filename)
            
            df = pd.DataFrame(all_data)
            if 'Quantity' in df.columns: df['Quantity'] = df['Quantity'].astype(str)
            df = df[df['Code'].fillna('').str.len() > 0]
            
            # Reorder
            cols = df.columns.tolist()
            if 'Source File' in cols:
                cols.remove('Source File')
                cols = ['Source File'] + cols
                df = df[cols]
                
            df.to_excel(output_path, index=False)
            
            self.status_var.set(f"Saved: {output_filename}")
            
            msg = f"Done! Processed {processed_count} files.\nSaved to: {output_filename}"
            if failed_files:
                msg += f"\n\nFailed: {', '.join(failed_files)}"
                
            messagebox.showinfo("Success", msg)

        except Exception as e:
            self.status_var.set("Error")
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()