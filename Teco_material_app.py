import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def select_source_file():
    file_path = filedialog.askopenfilename(title="Select Source File", filetypes=(("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")))
    entry_source.delete(0, tk.END)
    entry_source.insert(0, file_path)

def select_output_file():
    file_path = filedialog.asksaveasfilename(title="Select Output Location", defaultextension=".xlsx", filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
    entry_output.delete(0, tk.END)
    entry_output.insert(0, file_path)

def process_data():
    # Fetch filenames from the entry widgets
    file_name_original = entry_source.get()
    file_name_output = entry_output.get()

    # Determine the engine based on file extension
    if file_name_original.lower().endswith('.xls'):
        engine_to_use = 'xlrd'
    elif file_name_original.lower().endswith('.xlsx'):
        engine_to_use = 'openpyxl'
    else:
        messagebox.showerror("Error", "Unsupported file format.")
        return

    try:
        # 1. Load the data from Excel
        data = pd.read_excel(file_name_original, engine=engine_to_use)

        # 2. Extract relevant columns using 0-based column index
        subset_data = data.iloc[:, [7, 11, 6]]

        # 3. Rename columns
        subset_data.columns = ['Point', 'Span', 'FID']

        # 4. Filter unique rows
        unique_data = subset_data.drop_duplicates()

        # 5. Export/save the data back to Excel
        unique_data.to_excel(file_name_output, index=False)

        # Display success message
        messagebox.showinfo("Success", "Data processed and saved successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


# Create main tkinter window
root = tk.Tk()
root.title("Data Processing App")

# Create input fields, labels, and browse buttons
tk.Label(root, text="Source File Path").grid(row=0, column=0, padx=10, pady=10)
entry_source = tk.Entry(root, width=40)
entry_source.grid(row=0, column=1, padx=10, pady=10)
browse_source_btn = tk.Button(root, text="Browse", command=select_source_file)
browse_source_btn.grid(row=0, column=2, padx=10)

tk.Label(root, text="Output File Path").grid(row=1, column=0, padx=10, pady=10)
entry_output = tk.Entry(root, width=40)
entry_output.grid(row=1, column=1, padx=10, pady=10)
browse_output_btn = tk.Button(root, text="Browse", command=select_output_file)
browse_output_btn.grid(row=1, column=2, padx=10)

# Create and place the process button
process_button = tk.Button(root, text="Process Data", command=process_data)
process_button.grid(row=2, column=0, columnspan=3, pady=20)

root.mainloop()
