# Importing libraries
import sys
import csv
import pandas as pd
import pdfkit
import os
from datetime import date
import openpyxl


# Detect the delimiter of the csv file
def detect_csv_delimiter(file_path):
    with open(file_path, "r", newline="") as file:
        sample = file.read(4096)  # Read a sample of the file

        dialect = csv.Sniffer().sniff(sample)
        delimiter = dialect.delimiter

    return delimiter


# Read the csv file
def read_csv(file, delimiter):
    df = pd.read_csv(file, delimiter=delimiter)

    return df


# Extract Characteristics from the csv file
def extract_characteristics(att_col, primary_tokens):
    # Create a dictionary for the characteristics of all tokens
    all_characteristics = {}

    for primary_token in primary_tokens:
        # Get the rows with the same ProductPrimaryToken
        specific_rows = att_col[att_col["ProductPrimaryToken"] == primary_token]
        specific_df = pd.DataFrame(specific_rows)

        # Create a dictionary with the characteristics of the specific ProductPrimaryToken
        characteristics = {}

        # Iterate over the columns in the specific DataFrame
        for column in specific_df.columns:
            # Check the number of unique values in the column
            num_unique_values = specific_df[column].nunique()

            if num_unique_values == 1:
                characteristics[column] = specific_df[column].unique()[0]

        characteristics.pop("ProductPrimaryToken", None)
        all_characteristics[primary_token] = characteristics

    return all_characteristics


# Extract Attributes of a row
def extract_attributes(row, characteristics):
    attributes = {}

    for key, value in row.items():
        if key not in characteristics and not pd.isna(value) and value != "":
            attributes[key] = value

    return attributes


# Create a pdf file
def create_pdf(html_content, pdf_file):
    try:
        options = {
            "page-size": "A4",
            "margin-top": "3.67cm",
            "margin-right": "1.32cm",
            "margin-bottom": "2.54cm",
            "margin-left": "1.32cm",
            "encoding": "UTF-8",
            "quiet": "",
        }
        pdfkit.from_string(html_content, pdf_file, options=options)
    except Exception as e:
        print("Error while creating PDF:", str(e))


# Create a html file of the row data and return the html content
def generate_html(data):
    # Empty strings to store the <li> tags for characteristics, attributes and bullet points
    li_char_tags_1 = ""
    li_char_tags_2 = ""

    li_attr_tags_1 = ""
    li_attr_tags_2 = ""

    li_bullet_tags = ""

    counter = 0  # Counter to alternate between the two <div> tags

    # Iterate over the characteristics and attributes dictionaries to create the <li> tags
    for key, val in data["Characteristics"].items():
        key = key.replace("Attribute_", "")
        if counter % 2 == 0:
            li_char_tags_1 += f"<li>{key}: {val}</li>\n\t"
        else:
            li_char_tags_2 += f"<li>{key}: {val}</li>\n\t"
        counter += 1

    counter = 0
    for key, val in data["Attributes"].items():
        key = key.replace("Attribute_", "")
        if counter % 2 == 0:
            li_attr_tags_1 += f"<li>{key}: {val}</li>\n\t"
        else:
            li_attr_tags_2 += f"<li>{key}: {val}</li>\n\t"
        counter += 1

    # Split the bullet points by the new line character
    if isinstance(data["Attribute_BulletPointsProducto"], str):
        data["Attribute_BulletPointsProducto"] = data[
            "Attribute_BulletPointsProducto"
        ].split("\r\n")
        data["Attribute_BulletPointsProducto"] = [
            item.lstrip("- ") for item in data["Attribute_BulletPointsProducto"]
        ]
    elif isinstance(data["Attribute_BulletPointsProducto"], float):
        data["Attribute_BulletPointsProducto"] = str(
            data["Attribute_BulletPointsProducto"]
        ).split("\r\n")
        data["Attribute_BulletPointsProducto"] = [
            item.lstrip("- ") for item in data["Attribute_BulletPointsProducto"] if item
        ]

    # Create the <li> tags for the bullet points
    for item in data["Attribute_BulletPointsProducto"]:
        li_bullet_tags += f"<li>{item}</li>\n\t"

    # Css style for the html file
    style = """
     @import url("https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap");

  * {
    font-family: "Roboto", sans-serif;
    text-decoration: none;
  }

  :root {
    --primary: #0000002a;
    --secondary: #000000c7;
    --tertiary: #555555;
  }
  
  body {
    font-family: Arial, Helvetica, sans-serif;
  }

  .container {
    display: flex;
    flex-wrap: wrap;
    justify-content: space-between;
    align-items: flex-start;
    row-gap: 1rem;
    column-gap: 1rem;
    overflow: hidden;
  }

  .container img {
    width: 20rem;
    height: auto;
    max-height: 35rem;
    justify-self: center;
  }

  .info {
    display: flex;
    flex-direction: column;
    flex-basis: 50%;
  }

  .info p {
    margin-bottom: 1rem;
    color: var(--tertiary);
  }

  .info p span {
    font-weight: bold;
    color: var(--secondary);
  }

  .info :nth-child(3) {
    font-size: 20px;
  }

  .desc {
    background-color: var(--primary);
    color: #000000 !important;
    padding: 1rem;
    word-wrap: break-word;
  }

  .section {
    flex-basis: 100%;
    margin-top: 2rem;
  }

  #bullet {
    margin-top: 4rem;
  }

  .section ul{
    background-color: var(--primary);
    margin-top: 1rem;
    padding: 2rem;
	overflow: hidden;
  }

  .section div {
	width: 50%;
	float: left;
  }"""

    # Create the html content
    html_content = f"""
      <!DOCTYPE html>
  <html>
  <head>
      <title>{data['Token']}</title>
  </head>
  <style>{style}</style>
  <body>
      <div class="container">
        <picture>
            <source srcset="{data["Image_ProductPrimary"]}">
            <img src="{data["Image_ProductPrimary"]}">
        </picture>

          <div class="info">
              <p>Referencia de Producto: <span>{data["ProductPrimaryToken"]}</span></p>
              <p>Nombre de Producto: <span>{data["Name_es"]}</span></p>
              <p>Descripción de Producto:</p>
              <p class="desc">{data["ProductSection_T2_INFO_es"]}</p>
          </div>

          <section id="bullet" class="section">
              <h2>Bullet</h2>
              <ul>
                  {li_bullet_tags}
              </ul>
          </section>

          <section class="section">
              <h2>Caracteristicas</h2>
                <ul>
                    <div >
                        {li_char_tags_1}
                    </div>
                    <div >
                        {li_char_tags_2}
                    </div>
                </ul>
          </section>

          <section class="section">
              <h2>Atributos</h2>
                <ul>
                    <div >
                        {li_attr_tags_1}
                    </div>
                    <div >
                        {li_attr_tags_2}
                    </div>
                </ul>
          </section>
      </div>
  </body>
  </html>

      """

    with open("html_file.html", "w", encoding="utf-8") as file:
        file.write(html_content)

    return html_content


def main(file_path):
    # Get the file extension
    _, file_extension = os.path.splitext(file_path)

    # Read the csv or xlsx file
    if file_extension == ".csv":
        print("Reading " + file_path + "...")
        df = read_csv(file_path, detect_csv_delimiter(file_path)).sort_values(
            by="ProductPrimaryToken"
        )

    elif file_extension == ".xlsx":
        print("Reading " + file_path + "...")
        # Load the Excel file
        wb = openpyxl.load_workbook(file_path)
        sheet_name = wb.sheetnames[0]  # Assuming you want to read the first sheet
        sheet = wb[sheet_name]
        data = sheet.values
        df = pd.DataFrame(data)

        # Set column names if needed
        df.columns = df.iloc[0]

        # Convert all columns to string
        df = df.astype(str)

        # Sort DataFrame by 'ProductPrimaryToken' column
        df = df.sort_values(by="ProductPrimaryToken")

    # Drop columns with all NaN values and columns with all 0 values
    df = df.dropna(axis=1, how="all").drop(columns=df.columns[df.eq(0.0).all()])

    # Get columns 2, 3, and 4
    columns = df.iloc[:, [1, 2, 3, 4, 6, 8]]

    # Select columns that start with "Attribute_" (excluding "Attribute_BulletPointsProducto" and "Attribute_Estado")
    attribute_columns = df.filter(
        regex="^(?!Attribute_BulletPointsProducto|Attribute_Estado)(Attribute_)"
    ).copy()  # Create a copy of the filtered DataFrame

    # Add the ProductPrimaryToken column to the attribute_columns DataFrame
    attribute_columns["ProductPrimaryToken"] = df["ProductPrimaryToken"]

    # Get duplicate attribute columns based on ProductPrimaryToken
    duplicate_attribute_columns = attribute_columns[
        attribute_columns.duplicated(subset="ProductPrimaryToken", keep=False)
    ]

    # Get the unique tokens
    unique_tokens = columns["ProductPrimaryToken"].unique()

    # Create a dictionary to store the data of each row
    data = {
        "Token": "",
        "ProductPrimaryToken": "",
        "Name_es": "",
        "ProductSection_T2_INFO_es": "",
        "Image_ProductPrimary": "",
        "Attribute_BulletPointsProducto": "",
        "Characteristics": {},
        "Attributes": {},
    }

    print("Extracting characteristics for every PrimaryProductToken...")
    # Extract the characteristics and attributes of each row and create a pdf file
    chars = extract_characteristics(duplicate_attribute_columns, unique_tokens)

    # Iterate over the rows in the DataFrame and create a pdf file for each row
    for i in range(len(columns["ProductPrimaryToken"])):
        char = chars[
            columns["ProductPrimaryToken"][i]
        ]  # Get the characteristics of the row

        # Extract the attributes of the row
        attr = extract_attributes(attribute_columns.iloc[i].to_dict(), char)
        attr.pop("ProductPrimaryToken", None)

        # Add the data of the row to the data dictionary
        data["Token"] = columns["Token"][i]
        data["ProductPrimaryToken"] = columns["ProductPrimaryToken"][i]
        data["Name_es"] = columns["Name_es"][i]
        data["ProductSection_T2_INFO_es"] = columns["ProductSection_T2_INFO_es"][i]
        data["Image_ProductPrimary"] = columns["Image_ProductPrimary"][i]
        data["Attribute_BulletPointsProducto"] = columns[
            "Attribute_BulletPointsProducto"
        ][i]
        data["Characteristics"] = char
        data["Attributes"] = attr

        # Create the html content of the row
        html_content = generate_html(data)

        # Create the pdf file of the row and save it in the PDF folder of the current date folder
        if not os.path.exists("PDF"):
            print("Creating PDF folder...")
            os.makedirs("PDF")
        if not os.path.exists("PDF/" + str(date.today())):
            print("Creating " + str(date.today()) + " folder...")
            os.makedirs("PDF/" + str(date.today()))

        # Create the pdf file name
        pdf_file = data["ProductPrimaryToken"] + "_" + data["Token"] + ".pdf"

        # Create the pdf file path
        file_path = "PDF/" + str(date.today()) + "/" + pdf_file

        print("Creating " + pdf_file + "...")
        # Create the pdf file
        create_pdf(html_content, file_path)


if __name__ == "__main__":
    # Check if the file path is provided as an argument
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        main(file_path)
    else:
        print("Please provide the file path as an argument.")