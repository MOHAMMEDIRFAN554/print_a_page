import os
from tkinter import Tk, Label, Button, Entry, StringVar, IntVar, OptionMenu, messagebox
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, A5
import win32print
import win32api

# Generate the PDF file
def generate_pdf(file_name, pages, paper_size, page_type):
    width, height = paper_size
    c = canvas.Canvas(file_name, pagesize=paper_size)
    
    for _ in range(pages):
        if page_type == "Lined":
            # Draw lines for a lined page
            line_gap = 30
            y = height - 40
            while y > 40:
                c.line(40, y, width - 40, y)
                y -= line_gap
        c.showPage()
    
    c.save()

# Print PDF with or without duplex
def print_pdf(printer_name, file_name, duplex=False):
    hprinter = win32print.OpenPrinter(printer_name)
    printer_info = win32print.GetPrinter(hprinter, 2)
    devmode = printer_info['pDevMode']
    
    devmode.Duplex = 2 if duplex else 1
    
    win32print.ClosePrinter(hprinter)
    
    win32api.ShellExecute(
        0,
        "print",
        file_name,
        None,
        ".",
        0
    )

# GUI Application
class PrintingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Page Printing Tool")
        
        self.page_type = StringVar(value="Blank")
        self.pages = IntVar(value=1)
        self.paper_size = StringVar(value="A4")
        self.duplex = IntVar(value=0)
        self.selected_printer = StringVar()
        
        self.create_widgets()
    
    def create_widgets(self):
        Label(self.root, text="Select Page Type:").grid(row=0, column=0, padx=10, pady=5)
        OptionMenu(self.root, self.page_type, "Blank", "Lined").grid(row=0, column=1, padx=10, pady=5)
        
        Label(self.root, text="Enter Number of Pages:").grid(row=1, column=0, padx=10, pady=5)
        Entry(self.root, textvariable=self.pages).grid(row=1, column=1, padx=10, pady=5)
        
        Label(self.root, text="Select Paper Size:").grid(row=2, column=0, padx=10, pady=5)
        OptionMenu(self.root, self.paper_size, "A4", "A5").grid(row=2, column=1, padx=10, pady=5)
        
        Label(self.root, text="Enable Duplex Printing:").grid(row=3, column=0, padx=10, pady=5)
        OptionMenu(self.root, self.duplex, 0, 1).grid(row=3, column=1, padx=10, pady=5)
        
        Button(self.root, text="Print", command=self.start_printing).grid(row=4, column=0, columnspan=2, pady=10)
    
    def start_printing(self):
        # Get user input
        page_type = self.page_type.get()
        pages = self.pages.get()
        paper_size = A4 if self.paper_size.get() == "A4" else A5
        duplex = bool(self.duplex.get())
        
        if pages <= 0:
            messagebox.showerror("Error", "Number of pages must be greater than 0.")
            return
        
        # Generate PDF
        file_name = "pages_to_print.pdf"
        generate_pdf(file_name, pages, paper_size, page_type)
        
        # Get available printers
        printers = [printer[2] for printer in win32print.EnumPrinters(2)]
        if not printers:
            messagebox.showerror("Error", "No printers found.")
            return
        
        # Show printer selection
        self.selected_printer.set(printers[0])
        printer_choice = OptionMenu(self.root, self.selected_printer, *printers)
        printer_choice.grid(row=5, column=0, columnspan=2, pady=10)
        
        def send_to_printer():
            selected_printer = self.selected_printer.get()
            if not selected_printer:
                messagebox.showerror("Error", "Please select a printer.")
                return
            print_pdf(selected_printer, file_name, duplex)
            messagebox.showinfo("Success", "Printing initiated!")
        
        Button(self.root, text="Confirm Printer", command=send_to_printer).grid(row=6, column=0, columnspan=2, pady=10)

# Main Function
if __name__ == "__main__":
    root = Tk()
    app = PrintingApp(root)
    root.mainloop()
