import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfReader, PdfWriter
import fitz

def pdf_tool():
    selected_files = []

    def add_files():
        nonlocal selected_files
        new_files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if new_files:
            selected_files.extend(new_files)
            update_file_list()

    def remove_file():
        nonlocal selected_files
        selected_indices = file_list.curselection()
        for index in selected_indices[::-1]:
            del selected_files[index]
        update_file_list()

    def update_file_list():
        file_list.delete(0, tk.END)
        for index, file_path in enumerate(selected_files, start=1):
            file_list.insert(tk.END, f"{index}. {file_path}")

    def move_up():
        selected_index = file_list.curselection()
        if selected_index:
            index = selected_index[0]
            if index > 0:
                selected_files[index], selected_files[index - 1] = selected_files[index - 1], selected_files[index]
                update_file_list()
                file_list.selection_set(index - 1)

    def move_down():
        selected_index = file_list.curselection()
        if selected_index:
            index = selected_index[0]
            if index < len(selected_files) - 1:
                selected_files[index], selected_files[index + 1] = selected_files[index + 1], selected_files[index]
                update_file_list()
                file_list.selection_set(index + 1)

    def merge_pdfs():
        nonlocal selected_files
        if selected_files:
            pdf_writer = PdfWriter()
            for path in selected_files:
                pdf_reader = PdfReader(path)
                for page in pdf_reader.pages:
                    pdf_writer.add_page(page)
            save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if save_path:
                with open(save_path, 'wb') as output_pdf:
                    pdf_writer.write(output_pdf)
                messagebox.showinfo("Merge Complete", "PDFs merged successfully!")
        else:
            messagebox.showwarning("No Files Selected", "Please select PDF files to merge.")

        

    # Create the main window
    root = tk.Tk()
    root.title("PDF Merge Tool")

    # Create frame for file list and buttons
    file_frame = tk.Frame(root)
    file_frame.pack(padx=20, pady=10)

    # Create buttons
    add_button = tk.Button(file_frame, text="Add PDF Files", command=add_files)
    add_button.grid(row=0, column=0, padx=5)

    remove_button = tk.Button(file_frame, text="Remove Selected", command=remove_file)
    remove_button.grid(row=0, column=1, padx=5)

    move_up_button = tk.Button(file_frame, text="Move Up", command=move_up)
    move_up_button.grid(row=1, column=0, pady=5)

    move_down_button = tk.Button(file_frame, text="Move Down", command=move_down)
    move_down_button.grid(row=1, column=1, pady=5)

    merge_button = tk.Button(root, text="Merge PDFs", command=merge_pdfs)
    merge_button.pack(pady=10)

    # Create file listbox
    file_list = tk.Listbox(root, width=60, height=10)
    file_list.pack(padx=20, pady=5)

    # Start the main loop
    root.mainloop()

def pdf_split():
    # Your existing code for splitting PDFs
    selected_file_split = ""

    def browse_file():
        global selected_file_split
        selected_file_split = filedialog.askopenfilename(filetypes=[("PDF File", "*.pdf")])
        if selected_file_split:
            file_label.config(text=f"Selected File: {selected_file_split}")

    def validate_input(start_page, end_page, total_pages):
        try:
            start_page = int(start_page)
            end_page = int(end_page)
        except ValueError:
            messagebox.showwarning("Invalid Input", "Please enter valid page numbers.")
            return False

        if start_page < 1 or end_page > total_pages or start_page > end_page:
            messagebox.showwarning("Invalid Page Range", f"Please enter page numbers between 1 and {total_pages}.")
            return False

        return True

    def clear_fields():
        file_label.config(text="Selected File: None")
        start_page_entry.delete(0, tk.END)
        end_page_entry.delete(0, tk.END)
        global selected_file_split
        selected_file_split = ""

    def display_total_pages():
        if selected_file_split:
            try:
                pdf_reader = fitz.open(selected_file_split)
                total_pages = pdf_reader.page_count
                messagebox.showinfo("Total Pages", f"The selected PDF has {total_pages} pages.")
                pdf_reader.close()
            except Exception as e:
                messagebox.showwarning("Invalid PDF File", "Please select a valid PDF file.")
        else:
            messagebox.showwarning("No File Selected", "Please select a PDF file.")

    def save_split_pdf(pdf_viewer):
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if save_path:
            pdf_viewer.save(save_path)
            messagebox.showinfo("PDF Saved", "Split PDF saved successfully.")

    def show_selected_pages_viewer(start_page, end_page):
        if not selected_file_split:
            messagebox.showwarning("No File Selected", "Please select a PDF file.")
            return

        try:
            pdf_document = fitz.open(selected_file_split)
        except Exception as e:
            messagebox.showwarning("Invalid PDF File", "Please select a valid PDF file.")
            return

        start_page = int(start_page)
        end_page = int(end_page)

        if start_page < 1 or end_page > pdf_document.page_count or start_page > end_page:
            messagebox.showwarning("Invalid Page Range", f"Please enter page numbers between 1 and {pdf_document.page_count}.")
            return

        pdf_viewer = fitz.open()  # New PDF for selected pages

        for page_num in range(start_page - 1, end_page):
            pdf_viewer.insert_pdf(pdf_document, from_page=page_num, to_page=page_num)

        save_split_pdf(pdf_viewer)
        pdf_viewer.close()
        pdf_document.close()

        #messagebox.showinfo("Selected Pages Viewer", "Split PDF saved and opened in PDF viewer.")

    def split_pdf():
        start_page = start_page_entry.get()
        end_page = end_page_entry.get()

        if not selected_file_split:
            messagebox.showwarning("No File Selected", "Please select a PDF file.")
            return

        try:
            pdf_reader = fitz.open(selected_file_split)
            total_pages = pdf_reader.page_count
            pdf_reader.close()
        except Exception as e:
            messagebox.showwarning("Invalid PDF File", "Please select a valid PDF file.")
            return

        if validate_input(start_page, end_page, total_pages):
            start_page = int(start_page)
            end_page = int(end_page)
            show_selected_pages_viewer(start_page, end_page)

    root = tk.Tk()
    root.title("PDF Splitter Tool")

    file_frame = tk.Frame(root)
    file_frame.pack(padx=20, pady=10)

    select_file_button = tk.Button(file_frame, text="Select PDF File", command=browse_file)
    select_file_button.grid(row=0, column=0, padx=5)

    file_label = tk.Label(file_frame, text="Selected File: None")
    file_label.grid(row=0, column=1, padx=5)

    split_frame = tk.LabelFrame(root, text="Split Options")
    split_frame.pack(padx=20, pady=10)

    start_page_label = tk.Label(split_frame, text="Start Page:")
    start_page_label.grid(row=0, column=0, padx=5, pady=5)

    start_page_entry = tk.Entry(split_frame)
    start_page_entry.grid(row=0, column=1, padx=5, pady=5)

    end_page_label = tk.Label(split_frame, text="End Page:")
    end_page_label.grid(row=1, column=0, padx=5, pady=5)

    end_page_entry = tk.Entry(split_frame)
    end_page_entry.grid(row=1, column=1, padx=5, pady=5)

    split_button = tk.Button(root, text="Split PDF", command=split_pdf)
    split_button.pack(pady=10)

    clear_fields_button = tk.Button(root, text="Clear Fields", command=clear_fields)
    clear_fields_button.pack(pady=5)

    total_pages_button = tk.Button(root, text="Display Total Pages", command=display_total_pages)
    total_pages_button.pack(pady=5)

    root.mainloop()

        #**************************************************************************#

root = tk.Tk()
root.title("PDF Manipulator Tool")
root.geometry("300x270")  # Set window size to 300x300
root.resizable(False, False)  # Make window non-resizable

merge_button = tk.Button(root, text="Merge PDF",command=pdf_tool)
merge_button.pack(side=tk.LEFT, padx=40, pady=30)  # Place the Merge PDF button on the left

split_button = tk.Button(root, text="Split PDF",command=pdf_split)
split_button.pack(side=tk.LEFT, padx=10, pady=10)  # Place the Split PDF button next to Merge PDF button

root.mainloop()

