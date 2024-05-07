# Import necessary libraries
from fpdf import FPDF
import os
from datetime import datetime
import pandas as pd

# Read the Excel file
excel_path = "vpdata.xlsx"
df = pd.read_excel(excel_path)

# Create the names_dataset based on "BE05_01" and "BE05_02"
names_dataset = [f"{first} {last}" for first, last in zip(df["BE05_01"], df["BE05_02"])]

# Assuming df is your DataFrame containing the participant data
email_dataset = df["BE05_04"].tolist()

# Function to create PDF certificate
def create_certificate(name, email, output_folder):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=14)
    pdf.ln(1)
    pdf.image('uni.jpg', x=20, y=15, w=60)



     # Add two lines below (right-aligned at the top)
    pdf.set_y(15)
    pdf.set_font("Arial", size=11)
    pdf.cell(0, 10, txt="[PLACEHOLDER]", ln=True, align='R')
    pdf.cell(0, 5, txt="Psychologie", ln=True, align='R')

    # Add space between blocks
    pdf.ln(20)

    # Add lines of text at the top of the page
    pdf.set_font("Arial", style='B', size=14)
    pdf.cell(200, 10, txt="Nachweis über geleistet Versuchspersonstunden", ln=True, align='C')

    # Add space between blocks
    pdf.ln(3)

# Add additional text element with specified content
    pdf.set_font("Arial", 'B', size=10)  # Set font style to bold





# Shift the entire block 20 units to the right

    label_width = 40
    value_width = 110

    # Insert name from the function argument
    pdf.set_x(pdf.get_x() + 20)
    pdf.cell(label_width, 10, txt="Name Student/in:", ln=False, align='L')
    pdf.set_font("Arial", size=10)
    pdf.cell(value_width, 10, txt=name, ln=True, align='L')
    pdf.set_font("Arial", 'B', size=10)


    pdf.set_x(pdf.get_x() + 20) # shifts block to the right
    pdf.cell(label_width, 10, txt="VP-Stunden:", ln=False, align='L')
    pdf.set_font("Arial", size=10)  # Set font style to normal
    pdf.cell(value_width, 10, txt="1.0", ln=True, align='L')
    pdf.set_font("Arial", 'B', size=10)  # Set font style to normal

    pdf.set_x(pdf.get_x() + 20) # shifts block to the right
    pdf.cell(label_width, 10, txt="Name Dozent/in:", ln=False, align='L')
    pdf.set_font("Arial", size=10)  # Set font style to normal
    pdf.cell(value_width, 10, txt="[PLACEHOLDER]", ln=True, align='L')
    pdf.set_font("Arial", 'B', size=10)  # Set font style to normal

    pdf.set_x(pdf.get_x() + 20) # shifts block to the right
    pdf.cell(label_width, 10, txt="Titel der Studie:", ln=False, align='L')
    pdf.set_font("Arial", size=10)  # Set font style to normal
    pdf.multi_cell(value_width, 10, txt="[PLACEHOLDER]", align='L')
    pdf.set_font("Arial", 'B', size=10)  # Set font style to normal

    pdf.set_x(pdf.get_x() + 20) # shifts block to the right
    pdf.cell(label_width, 10, txt="Arbeitseinheit:", ln=False, align='L')
    pdf.set_font("Arial", size=10)  # Set font style to normal
    pdf.cell(value_width, 10, txt="[PLACEHOLDER]", ln=True, align='L')
    pdf.set_font("Arial", 'B', size=10)  # Set font style to normal

    pdf.set_x(pdf.get_x() + 20) # shifts block to the right
    pdf.cell(label_width, 10, txt="Datum:", ln=False, align='L')
    pdf.set_font("Arial", size=10)  # Set font style to normal
    pdf.cell(value_width, 10, txt=f"{datetime.now().strftime('%Y-%m-%d')}", ln=True, align='L')
    pdf.set_font("Arial", 'B', size=10)  # Set font style to normal

    pdf.set_x(pdf.get_x() + 20) # shifts block to the right
    pdf.cell(label_width, 10, txt="Unterschrift Dozent/in:", ln=False, align='L')
    pdf.set_font("Arial", size=10)  # Set font style to normal 
    pdf.cell(10)  # Adjust the value based on your preference   
# Add signature
    pdf.image('sig.jpg', x=pdf.get_x(), y=pdf.get_y(), w=60)
    pdf.ln(10)  # Move down after the image

    # Add space between blocks
    pdf.ln(10)

    # Reset font style to normal
    pdf.set_font("Arial", 'B', size=10)  # Set font style to bold for the right side



    # Create the output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)
    


    # Save the certificate with the individual's email in the specified output folder
    pdf.output(os.path.join(output_folder, f'{email}_certificate.pdf'))

# Specify the name of the subfolder to store certificates
output_subfolder = "certificates"
script_directory = os.path.dirname(os.path.realpath(__file__))

# Generate certificates in a loop
for name, email in zip(names_dataset, email_dataset):
    create_certificate(name, email, os.path.join(script_directory, output_subfolder))
