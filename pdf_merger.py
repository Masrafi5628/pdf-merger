import PyPDF2
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter import ttk
from PIL import Image, ImageTk
import fitz  # PyMuPDF for PDF rendering

class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Merger with Page Selection and Preview")
        self.pdf_files = []
        self.previews = []

        self.file_listbox = tk.Listbox(root, width=80, height=10)
        self.file_listbox.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

        self.add_button = tk.Button(root, text="Add PDFs", command=self.add_pdf)
        self.add_button.grid(row=1, column=0, padx=10, pady=10)

        self.merge_button = tk.Button(root, text="Merge PDFs", command=self.merge_pdfs)
        self.merge_button.grid(row=1, column=1, padx=10, pady=10)

        self.clear_button = tk.Button(root, text="Clear List", command=self.clear_list)
        self.clear_button.grid(row=1, column=2, padx=10, pady=10)

        self.preview_canvas = tk.Canvas(root, width=400, height=600, bg="gray")
        self.preview_canvas.grid(row=0, column=4, rowspan=10, padx=10, pady=10)

    def add_pdf(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        for file in files:
            start_page, end_page = self.get_page_range(file)
            if start_page is not None and end_page is not None:
                self.pdf_files.append((file, start_page, end_page))
                self.file_listbox.insert(tk.END, f"{file} (Pages: {start_page+1}-{end_page+1})")
                self.display_preview(file, start_page, end_page)

    def get_page_range(self, file):
        try:
            pdf = PyPDF2.PdfReader(file)
            total_pages = len(pdf.pages)

            start_page = simpledialog.askinteger("Page Range", f"Enter start page for {file} (1-{total_pages}):", initialvalue=1, minvalue=1, maxvalue=total_pages)
            end_page = simpledialog.askinteger("Page Range", f"Enter end page for {file} (1-{total_pages}):", initialvalue=total_pages, minvalue=start_page, maxvalue=total_pages)

            return start_page - 1, end_page - 1  # Adjust to 0-indexed
        except Exception as e:
            messagebox.showerror("Error", f"Error reading {file}: {str(e)}")
            return None, None

    def display_preview(self, file, start_page, end_page):
        try:
            doc = fitz.open(file)
            first_page = doc.load_page(start_page)
            last_page = doc.load_page(end_page)

            first_image = first_page.get_pixmap()
            last_image = last_page.get_pixmap()

            first_image = Image.open(fitz.BytesIO(first_image.tobytes()))
            last_image = Image.open(fitz.BytesIO(last_image.tobytes()))

            first_image = first_image.resize((200, 300))
            last_image = last_image.resize((200, 300))

            first_img = ImageTk.PhotoImage(first_image)
            last_img = ImageTk.PhotoImage(last_image)

            self.preview_canvas.create_image(100, 150, image=first_img)
            self.preview_canvas.create_image(300, 150, image=last_img)

            self.previews.append((first_img, last_img))
        except Exception as e:
            messagebox.showerror("Error", f"Error displaying preview: {str(e)}")

    def merge_pdfs(self):
        if not self.pdf_files:
            messagebox.showerror("Error", "No PDFs added.")
            return

        output_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if not output_file:
            return

        merger = PyPDF2.PdfWriter()

        try:
            for file, start_page, end_page in self.pdf_files:
                reader = PyPDF2.PdfReader(file)
                for page_num in range(start_page, end_page + 1):
                    merger.add_page(reader.pages[page_num])

            with open(output_file, 'wb') as output:
                merger.write(output)

            messagebox.showinfo("Success", "PDFs merged successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Error merging PDFs: {str(e)}")

    def clear_list(self):
        self.pdf_files.clear()
        self.file_listbox.delete(0, tk.END)
        self.preview_canvas.delete("all")
        self.previews.clear()


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFMergerApp(root)
    root.mainloop()
