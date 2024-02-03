# CSV_XLSX_TO_PDF

Python script to convert the CSV and XLSX files data to formatted pdfs

To use the script for the first time, follow these steps:

1.  Open the CSV_XLSX_TO_PDF folder in your Command Prompt (CMD) or Terminal (If you are using Linux)
2.  Write this line in the CMD/Terminal:  
    `pip install -r requirements.txt`
3.  Now run the pdf\*script.py using the following command:  
    `python pdf_script.py data.csv`

    If you are using Linux then:  
    `python3 pdf_script.py data.csv`

4.  Replace data.csv with the actual file path you want to process.

**Note:**

- The script will create a new folder named "PDFs" in the same directory where the script is located. The PDFs will be saved in this folder.
- If your are running this script on Ubuntu, make sure to install the following package using the following command:

  `sudo apt-get install wkhtmltopdf`
