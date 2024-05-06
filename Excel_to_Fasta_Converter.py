import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog

def excel_to_fasta(input_excel, output_fasta):
    # Read the Excel file using pandas
    df = pd.read_excel(input_excel)

    # Extract the base name of the input file without extension
    input_base_name = os.path.splitext(os.path.basename(input_excel))[0]

    # Open the output FASTA file for writing
    with open(output_fasta, 'w') as f:
        # Keep track of identifiers encountered
        identifier_count = {}
         
        # Iterate through the rows of the DataFrame
        for index, row in df.iterrows():
           
            sequence = str(row['Sequence from Start and End'])

            identifier = str(row['Input Sequence Identifier'])

            # Check if the identifier has been encountered before
            if identifier in identifier_count:
                # If yes, append a counter to the identifier
                identifier_count[identifier] += 1
                updated_identifier = f"{identifier}_{identifier_count[identifier]}"
            else:
                updated_identifier = identifier
                identifier_count[identifier] = 0

            # Write the identifier and sequence in FASTA format to the file
            f.write(f">{identifier}\n{sequence}\n")

    print(f"FASTA file '{output_fasta}' generated successfully.")
    return output_fasta

def browse_file():
    filename = filedialog.askopenfilename(title="Select Excel File")
    if filename:
        entry.delete(0, tk.END)
        entry.insert(0, filename)

def convert_to_fasta():
    input_excel_file = entry.get()
    if input_excel_file:
        output_file_name = os.path.splitext(os.path.basename(input_excel_file))[0]
        output_fasta_file = os.path.join('C:/Users/jyoth/Downloads', f"{output_file_name}.fasta")
        generated_file = excel_to_fasta(input_excel_file, output_fasta_file)
        result_label.config(text=f"FASTA file generated at: {generated_file}")
    else:
        result_label.config(text="Please select an input Excel file.")

# Create the main window
root = tk.Tk()
root.title("Excel to FASTA Converter")

# Create a label and entry for file path
tk.Label(root, text="Select Excel File:").pack(pady=5)
entry = tk.Entry(root, width=50)
entry.pack(pady=5, padx=10)
browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.pack(pady=5)

# Create a button to convert
convert_button = tk.Button(root, text="Convert to FASTA", command=convert_to_fasta)
convert_button.pack(pady=5)

# Create a label to display result
result_label = tk.Label(root, text="")
result_label.pack(pady=5)

# Create a disabled button
disabled_button = tk.Button(root, text="Developed in the Dept. of Biotechnology, SVCE", state=tk.DISABLED)
disabled_button.pack(pady=5)

root.mainloop()
