# Extract Bold, Italic, or Underlined Text from Excel 
This program extracts, or exports the **bold**, _italic_, or __underlined__ text from an Excel sheet. It assumes the 
sheet is named `ExtractText.xlsx` located in the same file as this program. It will then create three separate files 
with bold, italic, or underlined text respectively. In addition, it will 

## How to use?
1) Download and install Python.
2) Download this project locally via git, or unzip to its own folder.
3) Change directory to the folder on the command line.
4) Create a virtual environment.
   1) `python3 -m venv venv`
5) Activate the virtual environment.
   1) `source venv/bin/activate` on a mac.
   2) 'venv/Scripts/activate' on Windows.
6) Copy the spreadsheet you want to extract text from to this folder and name it `ExtractText.xlsx`.
7) Install the required packages.
   1) `pip install -r requirements.txt`
8) Run the program.
   1) `python3 extract_text.py`

## Important Libraries
The OpenPyXL library makes this program work. See [here](https://openpyxl.readthedocs.io/en/2.5.14/tutorial.html#loading-from-a-file)
for more details.

Real Python goes over in more details on how to use the OpenPyXL library: 
https://realpython.com/openpyxl-excel-spreadsheets-python/#practical-use-cases.

