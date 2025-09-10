# Certificate Generator

Automatically generate and send PDF certificates from Excel data.

## Features

- Generate PDF certificates from Excel files
- Send certificates via email
- Multiple certificate templates
- File organization
- Error handling and logging

## Requirements

Python 3.7+ with these packages:
```bash
pip install fpdf pandas openpyxl
```

## Quick Start

1. **Setup configuration:**
```bash
cp config_example.json config.json
```

2. **Edit config.json with your settings:**
```json
{
  "email": {
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587,
    "username": "your-email@gmail.com",
    "password": "your-app-password"
  },
  "certificate": {
    "title": "Certificate of Participation",
    "organization": "Your University",
    "instructor_name": "Dr. Your Name",
    "study_title": "Your Study Title"
  },
  "data": {
    "excel_file": "participant_data.xlsx",
    "name_first_column": "first_name",
    "name_last_column": "last_name",
    "email_column": "email_address"
  }
}
```

3. **Prepare Excel data with columns:**
   - First name
   - Last name
   - Email address

4. **Run the system:**
```bash
# Generate certificates
python main.py generate

# Send emails
python main.py send

# Run complete workflow
python main.py workflow
```

## Commands

```bash
python main.py generate              # Generate certificates
python main.py send                  # Send certificates via email
python main.py organize              # Organize files
python main.py workflow              # Run complete workflow
python main.py status                # Show status
python main.py test-email            # Test email configuration
```

## Examples

Create sample data:
```bash
python sample_data_template.py
```

Run with sample data:
```bash
python main.py workflow --excel sample_data/basic_participants.xlsx
```

## File Structure

```
certificate-generator/
├── main.py                    # Main script
├── config.py                  # Configuration
├── data_processor.py          # Data handling
├── certificate_generator.py   # PDF generation
├── email_sender.py           # Email sending
├── file_manager.py           # File organization
├── config.json               # Your settings
├── examples/                 # Usage examples
└── sample_data/              # Sample files
```

## Troubleshooting

- Check logs in `logs/certificate_generator.log`
- Verify Excel file has correct column names
- Test email with `python main.py test-email`
- Ensure image files (logo, signature) exist

## Customization

Create custom certificate templates by extending `CertificateTemplate` class in `certificate_generator.py`.

Create custom email templates using `EmailTemplate` class in `email_sender.py`.
