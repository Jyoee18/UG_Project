# AMR Sequence Retrieval

This Python script extracts antimicrobial resistance (AMR) sequences from a given FASTA file, using start and end positions provided in an Excel file. The extracted sequences, along with their lengths and associated resistance information, are saved to a new Excel file.

### Pre-requisites:
- **Python 3.x**
- Required packages:
  - `Biopython`
  - `openpyxl`

You can install the required packages using pip:
pip install biopython openpyxl

### Input Files:
1. **FASTA file**: Contains sequence data to be processed.
2. **Excel (.xlsx) file**: Contains columns for:
   - Sequence identifiers
   - Start and end positions for subsequences
   - Resistance information for each sequence.

### Output:
- A new Excel file will be generated with the following columns:
  - Sequence identifier
  - Full sequence
  - Sequence length
  - Start and end positions of the extracted subsequences
  - Extracted subsequence
  - Length of the extracted subsequence
  - Resistance information

### Usage:
To run the script, use the following command:
python amr_sequence_retrieval.py <input_fasta> <start_end_xlsx>

### Example:
python amr_sequence_retrieval.py sequences.fasta start_end_data.xlsx

### Error Handling:
- **FileNotFoundError**: If the input FASTA or Excel file is not found, an error message will be displayed.
- **ValueError**: If the start or end values cannot be converted to integers, an error message will be shown.
- **KeyError**: If a required column is missing in the Excel file, a warning will be issued.
- **General Exceptions**: Any other unexpected errors during file processing will display a generic error message.

### Script Overview:
1. **get_cut_sequence(sequence, start, end)**: Extracts a subsequence from a given sequence based on start and end positions.
2. **process_xlsx(xlsx_file)**: Reads the start, end, and resistance data from the Excel file and returns it as a dictionary.
3. **process_fasta(input_fasta, start_end_xlsx)**: Reads the input FASTA file, processes each sequence, and extracts the subsequences based on the data in the Excel file. The results are saved to a new Excel file.
4. **main()**: Handles command-line arguments, validates files, and calls the appropriate functions to process the data.

### Notes:
- The script checks if the provided FASTA and Excel files exist. If they don't, the script will terminate with an appropriate error message.
- The output Excel file will be saved in the same directory as the input FASTA file, with the name `input_identifier_AMR.xlsx`, where `input_identifier` is a part of the FASTA file name.

-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Excel to FASTA Converter

This script (`excel_to_fasta_converter.py`) converts sequence data from an Excel file into FASTA format. It appends unique identifiers to duplicated entries and saves the output in a FASTA file. Ideal for bioinformatics applications.

### Pre-requisites:
- **Python 3.x**
- Required packages: `pandas`
  
Install `pandas` by running:
pip install pandas

### Running the Script:
1. Place the `excel_to_fasta_converter.py` file in your desired directory.

2. Prepare the input Excel file (.xlsx) with two required columns:
   - `Input Sequence Identifier`: A unique identifier for each sequence.
   - `Sequence from Start and End`: The nucleotide or protein sequence to be converted.

3. In a command-line interface, navigate to the directory containing `excel_to_fasta_converter.py`:
   cd path/to/your/directory

4. Ensure the script is executable:
   chmod +x excel_to_fasta_converter.py
  
5. Run the script using:
   python excel_to_fasta_converter.py <input_excel_file> [output_fasta_file]
 
   - <input_excel_file>: Path to the Excel file.
   - [output_fasta_file]: (Optional) Path to save the FASTA file. If omitted, the output will be saved with the same name as the input file but with a `.fasta` extension.

### Example:
- To run with both input and output file names:
  python excel_to_fasta_converter.py sequences.xlsx output.fasta

- To run with only the input file:
  python excel_to_fasta_converter.py sequences.xlsx

  Output file will be `sequences.fasta`.

### Results:
- The FASTA file will contain each sequence and identifier in the following format:

  > Identifier
  Sequence

### Notes:
- Error handling:
  - `FileNotFoundError`: If the specified Excel file is not found.
  - `pd.errors.EmptyDataError`: If the Excel file is empty.
  - `ValueError`: If required columns are missing.






