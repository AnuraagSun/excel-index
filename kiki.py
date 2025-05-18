
import os
import re
import json
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import fitz  # PyMuPDF
import pandas as pd
from pathlib import Path

class TextbookIndexGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Textbook Index Generator")
        self.root.geometry("700x580")
        
        # Initialize variables
        self.pdf_path = tk.StringVar()
        self.output_folder = tk.StringVar()
        
        # Set default output folder to user's Documents
        documents_path = os.path.join(os.path.expanduser("~"), "Documents")
        self.output_folder.set(documents_path)
        
        # Create UI elements
        self.create_widgets()
        
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill=BOTH, expand=YES)
        
        # Application title
        title_label = ttk.Label(
            main_frame, 
            text="Textbook Index Generator", 
            font=("TkDefaultFont", 16, "bold")
        )
        title_label.pack(pady=10)
        
        # Description
        desc_label = ttk.Label(
            main_frame,
            text="Extract structured headings from PDF textbooks and create an Excel index",
            font=("TkDefaultFont", 10),
            wraplength=600
        )
        desc_label.pack(pady=(0, 20))
        
        # PDF File Selection
        file_frame = ttk.LabelFrame(main_frame, text="PDF Selection", padding=10)
        file_frame.pack(fill=X, pady=10)
        
        pdf_path_label = ttk.Label(file_frame, text="Selected File:")
        pdf_path_label.grid(row=0, column=0, sticky=W, pady=5)
        
        pdf_path_entry = ttk.Entry(file_frame, textvariable=self.pdf_path, width=50)
        pdf_path_entry.grid(row=0, column=1, padx=5, pady=5)
        
        pdf_browse_button = ttk.Button(
            file_frame, 
            text="Browse", 
            command=self.select_pdf_file,
            style="Accent.TButton"
        )
        pdf_browse_button.grid(row=0, column=2, padx=5, pady=5)
        
        # Output Folder Selection
        output_frame = ttk.LabelFrame(main_frame, text="Output Location", padding=10)
        output_frame.pack(fill=X, pady=10)
        
        output_path_label = ttk.Label(output_frame, text="Output Folder:")
        output_path_label.grid(row=0, column=0, sticky=W, pady=5)
        
        output_path_entry = ttk.Entry(output_frame, textvariable=self.output_folder, width=50)
        output_path_entry.grid(row=0, column=1, padx=5, pady=5)
        
        output_browse_button = ttk.Button(
            output_frame, 
            text="Browse", 
            command=self.select_output_folder,
            style="Accent.TButton"
        )
        output_browse_button.grid(row=0, column=2, padx=5, pady=5)
        
        # Lesson Mapping Frame
        mapping_frame = ttk.LabelFrame(main_frame, text="Lesson Index Mapping (JSON)", padding=10)
        mapping_frame.pack(fill=X, pady=10)
        
        # Example JSON
        example_json = '{\n  "1": "ELECTRIC CHARGES AND FIELDS",\n  "2": "COULOMB\'S LAW",\n  "3": "ELECTRIC FIELD"\n}'
        
        # JSON Text Area
        self.mapping_text = tk.Text(mapping_frame, height=10, width=80, wrap=tk.WORD)
        self.mapping_text.pack(fill=X, pady=5)
        self.mapping_text.insert(tk.END, example_json)
        
        # Add a scrollbar
        scrollbar = ttk.Scrollbar(mapping_frame, command=self.mapping_text.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.mapping_text.config(yscrollcommand=scrollbar.set)
        
        # Process Button Frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=X, pady=20)
        
        # Generate Button
        self.generate_button = ttk.Button(
            button_frame,
            text="Generate Excel Index",
            command=self.generate_index,
            style="success.TButton",
            width=25
        )
        self.generate_button.pack(padx=5, pady=5)
        
        # Status Label
        self.status_var = tk.StringVar()
        self.status_var.set("Ready to process...")
        self.status_label = ttk.Label(
            main_frame, 
            textvariable=self.status_var,
            font=("TkDefaultFont", 10, "italic")
        )
        self.status_label.pack(pady=10)
        
        # Progress Bar
        self.progress = ttk.Progressbar(
            main_frame, 
            orient="horizontal", 
            length=650, 
            mode="determinate"
        )
        self.progress.pack(fill=X, pady=5)
        
    def select_pdf_file(self):
        file_path = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            self.pdf_path.set(file_path)
            # Extract filename without extension as potential default output name
            pdf_name = os.path.splitext(os.path.basename(file_path))[0]
            self.status_var.set(f"Selected: {pdf_name}")
    
    def select_output_folder(self):
        folder_path = filedialog.askdirectory(
            title="Select Output Folder"
        )
        if folder_path:
            self.output_folder.set(folder_path)
    
    def update_status(self, message, progress_value=None):
        self.status_var.set(message)
        if progress_value is not None:
            self.progress["value"] = progress_value
        self.root.update_idletasks()
    
    def generate_index(self):
        # Validate inputs
        if not self.pdf_path.get():
            messagebox.showerror("Error", "Please select a PDF file.")
            return
        
        if not self.output_folder.get():
            messagebox.showerror("Error", "Please select an output folder.")
            return
        
        # Try to parse the lesson mapping JSON
        try:
            mapping_text = self.mapping_text.get("1.0", tk.END)
            lesson_mapping = json.loads(mapping_text)
            if not isinstance(lesson_mapping, dict):
                raise ValueError("Lesson mapping must be a dictionary")
        except json.JSONDecodeError:
            messagebox.showerror("Error", "Invalid JSON format in lesson mapping.")
            return
        except ValueError as e:
            messagebox.showerror("Error", str(e))
            return
        
        # Begin processing
        self.update_status("Starting PDF processing...", 10)
        self.generate_button.config(state=tk.DISABLED)
        
        try:
            # Extract text from PDF
            self.update_status("Extracting text from PDF...", 20)
            text_content = self.extract_text_from_pdf(self.pdf_path.get())
            
            # Parse headings
            self.update_status("Parsing headings...", 50)
            headings = self.extract_headings(text_content)
            
            # Create Excel file
            self.update_status("Creating Excel index...", 80)
            excel_path = self.create_excel_index(headings, lesson_mapping)
            
            # Complete
            self.update_status(f"Index generated successfully!", 100)
            messagebox.showinfo("Success", f"Excel index saved to:\n{excel_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.update_status(f"Error: {str(e)}", 0)
        
        self.generate_button.config(state=tk.NORMAL)
    
    def extract_text_from_pdf(self, pdf_path):
        """Extract text from PDF using PyMuPDF (fitz)"""
        text_content = ""
        try:
            doc = fitz.open(pdf_path)
            total_pages = doc.page_count
            
            for page_num in range(total_pages):
                # Update progress periodically
                if page_num % 10 == 0:
                    progress = 20 + int((page_num / total_pages) * 30)
                    self.update_status(f"Extracting page {page_num+1}/{total_pages}...", progress)
                
                page = doc[page_num]
                text_content += page.get_text()
                
            doc.close()
            return text_content
            
        except Exception as e:
            raise Exception(f"PDF extraction failed: {str(e)}")
    
    def extract_headings(self, text_content):
        """Extract structured headings using regex patterns"""
        headings = []
        
        # Split text into lines
        lines = text_content.split('\n')
        
        # Define regex patterns for different heading levels
        topic_pattern = re.compile(r'^(\d+\.\d+)\s+(.+)$')
        subtopic_pattern = re.compile(r'^(\d+\.\d+\.\d+)\s+(.+)$')
        sub_subtopic_pattern = re.compile(r'^(\d+\.\d+\.\d+\.\d+)\s+(.+)$')
        
        for line in lines:
            line = line.strip()
            
            # Check for sub-subtopic (most specific first to avoid partial matches)
            sub_subtopic_match = sub_subtopic_pattern.match(line)
            if sub_subtopic_match:
                number, title = sub_subtopic_match.groups()
                headings.append({
                    'level': 3,  # Level 3 = Sub-subtopic
                    'number': number,
                    'title': title.strip(),
                    'full_title': f"{number} {title.strip()}"
                })
                continue
                
            # Check for subtopic
            subtopic_match = subtopic_pattern.match(line)
            if subtopic_match:
                number, title = subtopic_match.groups()
                headings.append({
                    'level': 2,  # Level 2 = Subtopic
                    'number': number,
                    'title': title.strip(),
                    'full_title': f"{number} {title.strip()}"
                })
                continue
                
            # Check for topic
            topic_match = topic_pattern.match(line)
            if topic_match:
                number, title = topic_match.groups()
                headings.append({
                    'level': 1,  # Level 1 = Topic
                    'number': number,
                    'title': title.strip(),
                    'full_title': f"{number} {title.strip()}"
                })
        
        return headings
    
    def create_excel_index(self, headings, lesson_mapping):
        """Create an Excel index file with the specified format"""
        # Prepare data for DataFrame
        data = []
        
        for heading in headings:
            # Extract chapter number (first digit before the first dot)
            chapter_num = heading['number'].split('.')[0]
            
            # Get lesson name from mapping
            lesson_name = lesson_mapping.get(chapter_num, f"Chapter {chapter_num}")
            
            # Create a row with empty values
            row = {
                'Lesson name': '',
                'Topic': '',
                'Subtopic': '',
                'Sub-subtopic': ''
            }
            
            # Fill in the appropriate level
            if heading['level'] == 1:  # Topic
                row['Lesson name'] = lesson_name
                row['Topic'] = heading['full_title']
            elif heading['level'] == 2:  # Subtopic
                row['Subtopic'] = heading['full_title']
            elif heading['level'] == 3:  # Sub-subtopic
                row['Sub-subtopic'] = heading['full_title']
                
            data.append(row)
        
        # Create DataFrame
        df = pd.DataFrame(data)
        
        # Create Excel filename based on the PDF name
        pdf_name = os.path.splitext(os.path.basename(self.pdf_path.get()))[0]
        excel_filename = f"{pdf_name}_index.xlsx"
        excel_path = os.path.join(self.output_folder.get(), excel_filename)
        
        # Write to Excel
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Textbook Index")
            
            # Get the workbook and the worksheet
            workbook = writer.book
            worksheet = writer.sheets["Textbook Index"]
            
            # Format the worksheet
            # Set column widths
            worksheet.column_dimensions['A'].width = 30  # Lesson name
            worksheet.column_dimensions['B'].width = 40  # Topic
            worksheet.column_dimensions['C'].width = 40  # Subtopic
            worksheet.column_dimensions['D'].width = 40  # Sub-subtopic
            
        return excel_path


if __name__ == "__main__":
    # Create the main window with a Bootstrap style
    root = ttk.Window(themename="litera")  # You can choose other themes like: cosmo, cyborg, darkly, etc.
    app = TextbookIndexGenerator(root)
    root.mainloop()
