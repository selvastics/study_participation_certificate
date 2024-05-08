# Create Study Participant Certificate
## Overview
This repository contains scripts to automate the creation and distribution of PDF certificates for study participants. I used it to automatically credit students for the time they spent participating in one of my studies, which is a partial requirement for their B.Sc. degree. The script processes participant data from an Excel file, generates personalized certificates, emails them, and organizes the sent files into a designated folder. It should be quite intuitive to adapt the code to a user's specific needs. I also provide an XML file for SoSci Survey (https://www.soscisurvey.de) implementation, which works with the default code. Feel free to contact me if you need any assistance. 

## Requirements
- Python 3.x
- Libraries: fpdf, pandas
  ```bash
  pip install fpdf pandas
  ```

## File Structure
- **Data File**: `vpdata.xlsx` contains participant details with columns for first name, last name, ID number, and email.
- **Certificate Assets**:
  - `uni.jpg`: University logo.
  - `sig.jpg`: Signature and stamp image.

## Scripts
- **write.py**: Generates PDF certificates and saves them to the `certificates` folder.
- **send.py**: Emails the certificates. Requires email server configuration and user credentials.
- **move.py**: Moves sent certificates to another folder. Requires updating file names in the script.

## Usage
Run the following commands in your terminal:
```bash
python write.py  # Generates PDFs
python send.py   # Sends emails
python move.py   # Moves sent PDFs
```

## Configuration
- **send.py**: Set your email server details (e.g., SMTP server, port) and authentication credentials.
- **move.py**: List the filenames in the `certificates` folder that you want to move.

## Customization
You may need to adjust the text in the PDF templates or the email body to fit your specific study requirements.

## Additional Information
Ensure the `uni.jpg` and `sig.jpg` files are in the correct directory as specified in the scripts. Modify the paths if necessary.


![Alt text](/example.png)
