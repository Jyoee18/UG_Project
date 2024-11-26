# -*- coding: utf-8 -*-
"""
@author: jyothishree

Script Name: excel_to_fasta_converter.py

Version: 3.00

Description:
This script converts sequence data from an Excel file into FASTA format. Each sequence in the Excel file is given an identifier, and any duplicate identifiers are appended with an incremented counter to ensure uniqueness in the output FASTA file. The resulting FASTA file is saved to the specified output path.

The script requires an input Excel (.xlsx) file and an output file path for the generated FASTA file. 

Input:
    1. An Excel file (.xlsx) containing two required columns:
       - 'Input Sequence Identifier': A unique identifier for each sequence.
       - 'Sequence from Start and End': The nucleotide or protein sequence to convert to FASTA format.

Output:
    A FASTA file containing sequences and identifiers in the following format:
    >Identifier
    Sequence

Usage:
    python excel_to_fasta_converter.py <input_excel> <output_fasta>

Example:
    python excel_to_fasta_converter.py sequences.xlsx output.fasta

Error Handling:
    1. FileNotFoundError: Displays an error message if the input Excel file is not found.
    2. pd.errors.EmptyDataError: Shows an error if the Excel file is empty.
    3. ValueError: If required columns ('Input Sequence Identifier' and 'Sequence from Start and End') are missing from the Excel file.
    4. General Exception: Any unexpected errors are captured and displayed to the user for debugging purposes.
"""
# Importing necessary modules
import pandas as pd  
import os  
import sys  

def excel_to_fasta(input_excel, output_fasta):
    """
    Converts an Excel file containing sequences into a FASTA file.
    
    Parameters:
    input_excel (str): Path to the input Excel file.
    output_fasta (str): Path to the output FASTA file.

    Returns:
    str: Path to the generated FASTA file.
    """
    try:
        # Read the Excel file using pandas
        df = pd.read_excel(input_excel)
        
        # Check if necessary columns exist
        if 'Sequence from Start and End' not in df.columns or 'Input Sequence Identifier' not in df.columns:
            raise ValueError("Excel file must contain 'Sequence from Start and End' and 'Input Sequence Identifier' columns.")
        
        # Open the output FASTA file for writing
        with open(output_fasta, 'w') as f:
            # Keep track of identifiers encountered to avoid duplicates
            identifier_count = {}

            # Iterate through each row in the DataFrame
            for index, row in df.iterrows():
                # Ensure the sequence is a string
                sequence = str(row['Sequence from Start and End'])
                
                # Retrieve identifier and ensure it is a string
                identifier = str(row['Input Sequence Identifier'])

                # Handle duplicate identifiers by appending a counter
                if identifier in identifier_count:
                    identifier_count[identifier] += 1
                    updated_identifier = f"{identifier}_{identifier_count[identifier]}"
                else:
                    updated_identifier = identifier
                    identifier_count[identifier] = 0

                # Write the identifier and sequence in FASTA format
                f.write(f">{updated_identifier}\n{sequence}\n")

        print(f"FASTA file '{output_fasta}' generated successfully.")
        return output_fasta
    
    except FileNotFoundError:
        print(f"Error: The file '{input_excel}' was not found.")
    except pd.errors.EmptyDataError:
        print("Error: The Excel file is empty.")
    except ValueError as ve:
        print(f"Error: {ve}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

def main():
    # Check for at least one argument
    if len(sys.argv) < 2:
        print("Usage: python script.py <input_excel_file> [output_fasta_file]")
        sys.exit(1)

    # Get input Excel file from command-line arguments
    input_excel_file = sys.argv[1]

    # Determine output FASTA file name
    if len(sys.argv) == 3:
        output_fasta_file = sys.argv[2]
    else:
        # If only input file is provided, generate output filename based on input
        output_base_name = os.path.splitext(input_excel_file)[0]
        output_fasta_file = f"{output_base_name}.fasta"

    # Run the conversion function
    excel_to_fasta(input_excel_file, output_fasta_file)

if __name__ == "__main__":
    main()
