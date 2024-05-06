"# Bioinformatics-Python-Scripts" 

**AMR Sequence Retrieval Tool**

## Description

The AMR (Antimicrobial Resistance) Sequence Retrieval Tool is a Python script that facilitates the retrieval of AMR-related sequences from a given input FASTA file based on start and end positions provided in an Excel file. It processes the input files and generates an Excel file containing the retrieved sequences along with their associated resistance information.

## Features

- Retrieve AMR-related sequences from a given input FASTA file.
- Utilize start and end positions provided in an Excel file to extract sequences.
- Generate an Excel file with the retrieved sequences and their resistance information.

## Dependencies

- Python 3.x
- Biopython
- openpyxl
- tkinter (included in standard Python library)

## Installation

1. Clone this repository to your local machine.
2. Install the required dependencies using pip: pip install biopython openpyxl
   
## Usage

1. Run the script using Python: python amr_gui.py
2. Browse and select the input FASTA file containing the sequences you want to process.
3. Browse and select the Excel file containing the start and end positions for the sequences.
4. Click on the "Retrieve AMR Sequences" button to initiate the sequence retrieval process.
5. The output Excel file with the retrieved sequences will be generated in the Downloads directory.
-----------------------------------------------------------------------------------------------------------------------------
**Excel to FASTA Converter**

## Description

The Excel to FASTA Converter is a Python script that allows users to convert sequences stored in an Excel file into a FASTA file format commonly used in bioinformatics applications. This tool simplifies the process of converting data between these two formats, making it easier to work with sequence data in various bioinformatics workflows.

## Features

- Conversion of Excel to FASTA: The converter allows users to convert sequence data stored in an Excel file into the FASTA file format commonly used in bioinformatics.
- User-friendly Interface: The tool provides a simple and intuitive graphical user interface (GUI) using tkinter, making it easy for users to select input files and initiate the conversion process.
- Automatic Identifier Handling: The converter automatically handles identifiers in the input Excel file, ensuring that each sequence in the output FASTA file has a unique identifier.
- Output Location: The converted FASTA file is generated in the default download directory with the same base name as the input Excel file, simplifying the process of locating the output file.

## Dependencies

- pandas
- tkinter

## Installation

1. Clone this repository to your local machine.
2. Install the required dependencies using pip: pip install pandas

## Usage

1. **Select Excel File**: Click the "Browse" button to select the input Excel file containing the sequences to be converted.

2. **Convert to FASTA**: After selecting the input Excel file, click the "Convert to FASTA" button to initiate the conversion process.

3. **Result**: The converted FASTA file will be generated in the default download directory with the same base name as the input Excel file.





