# import os 
import tabula 
import camelot
import html
import tkinter as tk
import pandas as pd
from tkinter import filedialog, messagebox, scrolledtext

# JAVA_PATH = r"C:\Program Files\Eclipse Adoptium\jdk-21.0.8.9-hotspot\bin"
# os.environ['JAVA_HOME'] = JAVA_PATH
# os.environ['PATH'] = JAVA_PATH + ";" + os.environ['PATH']

def convert_to_excel(file_path, output_path ):
        try:
        # Direct convert (may create multiple sheets)
         tables = camelot.read_pdf(file_path, pages="all", flavor="stream")
         df_all = pd.concat([t.df for t in tables], ignore_index=True)
         df_all.to_excel(output_path, index=False)
         messagebox.showinfo("Success", f"Excel file saved at:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
def upload_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("PDF files", "*.pdf")]
    )
    if file_path:

    # Ask user where to save
       output_path = filedialog.asksaveasfilename(
         defaultextension=".xlsx",
         filetypes=[("Excel files", "*.xlsx"), ("Excel files", "*.xls")]
       )
       if output_path:
           convert_to_excel(file_path, output_path)
            


# GUI setup
root = tk.Tk()
root.title("Pdf to XML Converter")
root.geometry("600x400")

label = tk.Label(root, text="Upload PDf Bank Statement", font=("Arial", 14))
label.pack(pady=10)

upload_btn = tk.Button(root, text="Choose File & Save As Excel", command=upload_file, font=("Arial", 12))
upload_btn.pack(pady=5)

# save_btn = tk.Button(root, text="Save XML", command=save_file, font=("Arial", 12))
# save_btn.pack(pady=5)

# Text area to show XML

root.mainloop()
