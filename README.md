"# Bioinformatics-Python-Scripts" 

**AMR Sequence Retrieval Tool**

**Description**

The AMR (Antimicrobial Resistance) Sequence Retrieval Tool is a Python script that facilitates the retrieval of AMR-related sequences from a given input FASTA file based on start and end positions provided in an Excel file. It processes the input files and generates an Excel file containing the retrieved sequences along with their associated resistance information.

**Features**

1.Retrieve AMR sequences from a given input FASTA file.

2.Utilize start and end positions provided in an Excel file to extract sequences.

3.Generate an Excel file with the retrieved sequences and their resistance information.

**Dependencies**

Python 3.x, 
Biopython, 
openpyxl, 
tkinter (included in standard Python library)

**Installation**

1.Clone this repository to your local machine.

2.Install the required dependencies using pip: pip install biopython openpyxl

**Usage**

1.Run the script using Python: python amr_gui.py

2.Browse and select the input FASTA file containing the sequences you want to process.

3.Browse and select the Excel file containing the start and end positions for the sequences.

4.Click on the "Retrieve AMR Sequences" button to initiate the sequence retrieval process.

5.The output Excel file with the retrieved sequences will be generated in the Downloads directory.



