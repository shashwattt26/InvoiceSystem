import os
import subprocess
import datetime as dt
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

import docx


class InvoiceAutomation:
    def __init__(self):
        self.root = tk.Tk() 
        self.root.title("Invoice Automation")
        self.root.geometry("500x600")

        self.partner_date_label = tk.label(self.root, text="Date")
        self.partner_name_label = tk.label(self.root, text="Name")
        self.partner_add_label = tk.label(self.root, text="Address")
        self.partner_cont_label = tk.label(self.root, text="Contact Number")
        self.partner_invoiceNo_label = tk.label(self.root, text="inumber")
        self.partner_service_label = tk.label(self.root, text="Service")
        self.partner_amount_label = tk.label(self.root, text="Amount")
        self.partner_price_label = tk.label(self.root, text="Price")
        self.partner_payment_method_label = tk.label(self.root, text="Payment Method")

        self.payment_methode = {
            "SBI": {
                "Recipient": "AmanInfotech",
                "Bank": "State Bank of India",
                "AccNo": "1234567890",
                "ISFC": "SBIX0000001"
                },
            "HDFC": {
                "Recipient": "AmanInfotech",
                "Bank": "HDFC Bank",
                "AccNO": "2345678901",
                "ISFC": "HDFCX0000001"
                },
            "ICICI": {
                "Recipient": "AmanInfotech",
                "Bank": "ICICI Bank",
                "AccNo": "3456789012",
                "ISFC": "ICICX0000001"
                }
            }
        self.partner_date_entry = tk.Entry(self.root)
        self.partner_name_entry = tk.Entry(self.root)
        self.partner_add_entry = tk.Entry(self.root)
        self.partner_cont_entry = tk.Entry(self.root)
        self.partner_invoiceNo_entry = tk.Entry(self.root)
        self.partner_service_entry = tk.Entry(self.root)
        self.partner_amount_entry = tk.Entry(self.root)
        self.partner_price_entry = tk.Entry(self.root)

        self.partner_payment_method = tk.StringVar(self.root)
        self.partner_payment_method.set('SBI')

        self.payment_method_dropdown = tk.OptionMenu(self.root, self.payment_method, *self.payment_methode.keys())

        self.create_button = tk.Button(self.root, text="Create Invoice", command=self.create_invoice)

        padding_options = {"fill": "x" , "expand" : True, "padx": 5, "pady": 2}

        self.partner_date_label.pack(padding_options)
        self.partner_date_entry.pack(padding_options)

        self.partner_name_label.pack(padding_options)
        self.partner_name_entry.pack(padding_options)

        self.partner_add_label.pack(padding_options)
        self.partner_add_entry.pack(padding_options)

        self.partner_cont_label.pack(padding_options)
        self.partner_cont_entry.pack(padding_options)

        self.partner_invoiceNo_label.pack(padding_options)
        self.partner_invoiceNo_entry.pack(padding_options)

        self.partner_service_label.pack(padding_options)
        self.partner_service_entry.pack(padding_options)

        self.partner_amount_label.pack(padding_options)
        self.partner_amount_entry.pack(padding_options)

        self.partner_price_label.pack(padding_options)
        self.partner_price_entry.pack(padding_options)

        self.partner_payment_method_label.pack(padding_options)
        self.payment_method_dropdown.pack(padding_options)

        self.create_button.pack(padding_options)

        self.root.mainloop()
        

    def create_invoice(self):
        pass

if __name__ == "__main__":
    InvoiceAutomation()