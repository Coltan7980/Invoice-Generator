import os
import sys
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
from Scripts.GrabbingData import GrabPartData, GetParts, GetSubtotalTaxAndTotalDue
from docxtpl import DocxTemplate

def GetTemplatePath():
    global templatePath 
    templatePath = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    with open(os.path.join(os.getcwd(), "data") + "\\template.txt", "w") as f:
        f.write(templatePath + "\n")
    
def OpenFile():
    global pdfPath
    pdfPath = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    

def GetDirectory():
    global directoryPath
    
    folder_path = os.path.join(os.getcwd(), "data")
    os.makedirs(folder_path, exist_ok=True)
    file_path = os.path.join(folder_path, "directory.txt")
    directoryPath = filedialog.askdirectory()
    with open(file_path, "a+") as f:
        f.write(directoryPath + "\n")

def DirectoryLocation():
    with open(os.path.join(os.getcwd(), "data") + "\\directory.txt", "r") as f:
        lines = f.readlines()
        if lines:
            return lines[-1].strip()
        else:
            return None

def TemplateLocation():
    if getattr(sys, 'frozen', False):
        basePath = sys._MEIPASS
    else:
        basePath = os.path.dirname(__file__)
    return os.path.join(basePath, "Templates", "Template2025.docx") 


    
    

def GenerateInvoice():
    if templatePath and pdfPath:
        data = GrabPartData(pdfPath)
        parts = GetParts(data)
        dic = GetSubtotalTaxAndTotalDue(pdfPath)
        # Handle the data in the dictionary
        subtotal = f"{float(dic['Subtotal'].strip('$')):.2f}"
        tax = f"{float(dic['Tax'].strip('$')):.2f}"
        totalDue = f"{float(dic['TotalDue'].strip('$')):.2f}"
        #Handle the parts data
        customer_name = entry_name.get()
        shipping_cost = f"{float(entry_shipping.get()):.2f}"
        labor_desc = entry_labor_desc.get()
        labor_hours = entry_labor_hours.get()
        labor_price = f"{float(entry_labor_price.get()):.2f}"
        grandTotal = f"{float(totalDue) + float(shipping_cost) + float(labor_price):.2f}"
        
        context = {
            'customerName': customer_name,
            'shippingCost': shipping_cost,
            'laborDesc': labor_desc,
            'laborHours': labor_hours,
            'laborCost': labor_price,
            'subtotal': subtotal,
            'tax': tax,
            'grandTotal': grandTotal,
            'parts': parts,
            'notes' : entry_notes.get(),
        }
        doc = DocxTemplate(templatePath)
        doc.render(context)
        doc.save(f"{directoryPath}\\Invoice_{customer_name}.docx")
        messagebox.showinfo("Success", f"Invoice generated successfully at {directoryPath}Invoice_{customer_name}.docx")
    else:
        messagebox.showerror("Error", "Please select both a PDF file and a template.")
        return
    


    
if __name__ == "__main__":

    window = tk.Tk()
    window.title("Invoice Generator")
    window.geometry("1920x1080")
    tk.Label(window, text="Customer Name:").grid(row=0, column=0, sticky="e")
    tk.Label(window, text="Shipping Cost:").grid(row=1, column=0, sticky="e")
    tk.Label(window, text="Labor Description:").grid(row=2, column=0, sticky="e")
    tk.Label(window, text="Labor Hours:").grid(row=3, column=0, sticky="e")
    tk.Label(window, text="Labor Price:").grid(row=4, column=0, sticky="e")
    tk.Label(window, text="Notes:").grid(row=5, column=0, sticky="e")

    button = tk.Button(window, text="Select PDF", command=OpenFile)
    button.grid(row=8, column=0, columnspan=2, pady=10)
    button = tk.Button(window, text="Select Template", command=GetTemplatePath)
    button.grid(row=9, column=0, columnspan=2, pady=10)
    button = tk.Button(window, text="choose directory", command=GetDirectory)
    button.grid(row=10, column=0, columnspan=2, pady=10)

    entry_name = tk.Entry(window, width=30)
    entry_shipping = tk.Entry(window, width=30)
    entry_labor_desc = tk.Entry(window, width=30)
    entry_labor_hours = tk.Entry(window, width=30)
    entry_labor_price = tk.Entry(window, width=30)
    entry_notes = tk.Entry(window, width=30)
    templatePath = None
    pdfPath = None
    try:
        directoryPath = DirectoryLocation()
    except FileNotFoundError:
        directoryPath = None

    try:
        templatePath = TemplateLocation()
    except FileNotFoundError:
        messagebox.showerror("Error", "Template file not found. Please select a template.")




    entry_name.grid(row=0, column=1)
    entry_shipping.grid(row=1, column=1)
    entry_labor_desc.grid(row=2, column=1)
    entry_labor_hours.grid(row=3, column=1)
    entry_labor_price.grid(row=4, column=1)
    entry_notes.grid(row=5, column=1)

    tk.Button(window, text="Generate Invoice", command=GenerateInvoice).grid(row=6, column=0, columnspan=2, pady=10)
    window.mainloop() 