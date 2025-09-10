# Certificate Generator

A comprehensive, flexible, and easy-to-use system for automatically generating and distributing PDF certificates for study participants, research volunteers, course completions, and more.

## 🌟 Features

- **Flexible Data Input**: Works with Excel files in various formats
- **Multiple Certificate Templates**: Default, study participation, and custom templates
- **Automated Email Distribution**: Send certificates via email with customizable templates
- **File Organization**: Automatic organization of sent and unsent certificates
- **Comprehensive Logging**: Detailed logging for troubleshooting and monitoring
- **Configuration Management**: Easy configuration through JSON files
- **Error Handling**: Robust error handling and validation
- **Extensible Design**: Easy to customize and extend for specific needs

## 📋 Requirements

- Python 3.7+
- Required libraries: `fpdf`, `pandas`, `openpyxl`

Install dependencies:
```bash
pip install fpdf pandas openpyxl
```

## 🚀 Quick Start

### 1. Basic Usage

```bash
# Generate certificates from Excel data
python main.py generate

# Send certificates via email
python main.py send

# Organize files (move sent certificates)
python main.py organize

# Run complete workflow
python main.py workflow
```

### 2. Configuration

Copy the example configuration and customize it:

```bash
cp config_example.json config.json
```

Edit `config.json` with your settings:

```json
{
  "email": {
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587,
    "username": "your-email@gmail.com",
    "password": "your-app-password",
    "sender_name": "Research Team"
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

### 3. Prepare Your Data

Create an Excel file with participant data. Required columns:
- First name
- Last name  
- Email address

Optional columns:
- Participant ID
- Any additional data you want to include

## 📖 Detailed Usage

### Command Line Interface

```bash
# Generate certificates with specific template
python main.py generate --template study_participation --excel data.xlsx

# Send emails with specific template
python main.py send --template german

# Test email configuration
python main.py test-email

# Show application status
python main.py status

# Run complete workflow
python main.py workflow --cert-template study_participation --email-template german
```

### Programmatic Usage

```python
from data_processor import DataProcessor
from certificate_generator import CertificateGenerator
from email_sender import EmailSender

# Process data
data_processor = DataProcessor()
data_processor.load_data("participants.xlsx")
data_processor.process_data()
participants = data_processor.get_participants()

# Generate certificates
cert_generator = CertificateGenerator("study_participation")
results = cert_generator.generate_certificates(participants)

# Send emails
email_sender = EmailSender()
email_sender.send_bulk_certificates(participants, "certificates/", "german")
```

## 🎨 Customization

### Custom Certificate Templates

Create your own certificate template:

```python
from certificate_generator import CertificateTemplate

class MyCustomTemplate(CertificateTemplate):
    def create_certificate(self, pdf, participant):
        # Your custom certificate design
        pdf.set_font("Arial", size=16)
        pdf.cell(200, 10, txt="My Custom Certificate", ln=True, align='C')
        # ... add more custom elements
```

### Custom Email Templates

```python
from email_sender import EmailSender, EmailTemplate

email_sender = EmailSender()

# Add custom template
custom_template = EmailTemplate(
    subject="Your Certificate - {study_title}",
    body="Dear {participant_name}, your certificate is attached!",
    is_html=False
)
email_sender.add_template("custom", custom_template)
```

## 📁 File Structure

```
certificate-generator/
├── main.py                    # Main orchestration script
├── config.py                  # Configuration management
├── logger.py                  # Logging system
├── data_processor.py          # Data processing and validation
├── certificate_generator.py   # PDF certificate generation
├── email_sender.py           # Email sending functionality
├── file_manager.py           # File organization
├── config.json               # Configuration file
├── config_example.json       # Example configuration
├── examples/                 # Usage examples
│   ├── basic_usage.py
│   ├── custom_template.py
│   └── email_templates.py
├── sample_data/              # Sample data files
├── certificates/             # Generated certificates
├── sended/                   # Sent certificates
└── logs/                     # Log files
```

## 🔧 Configuration Options

### Email Configuration
- `smtp_server`: SMTP server address
- `smtp_port`: SMTP port (usually 587 for TLS, 465 for SSL)
- `username`: Email username
- `password`: Email password or app password
- `use_tls`: Use TLS encryption
- `use_ssl`: Use SSL encryption
- `sender_name`: Display name for sender
- `sender_email`: Sender email address

### Certificate Configuration
- `title`: Certificate title
- `subtitle`: Certificate subtitle
- `organization`: Organization name
- `department`: Department name
- `instructor_name`: Instructor/Principal name
- `study_title`: Study or course title
- `work_unit`: Work unit or department
- `hours`: Hours or credits
- `logo_path`: Path to logo image
- `signature_path`: Path to signature image
- `output_folder`: Folder for generated certificates
- `sent_folder`: Folder for sent certificates

### Data Configuration
- `excel_file`: Path to Excel data file
- `name_first_column`: Column name for first name
- `name_last_column`: Column name for last name
- `email_column`: Column name for email address
- `id_column`: Column name for participant ID (optional)

## 📊 Examples

### Example 1: Basic Study Participation

```bash
# 1. Prepare your data in Excel with columns: first_name, last_name, email
# 2. Configure settings in config.json
# 3. Run the workflow
python main.py workflow
```

### Example 2: German University Study

```bash
# Use the German template for university studies
python main.py workflow --cert-template study_participation --email-template german
```

### Example 3: Custom Template

```python
# Create custom template and use it
python examples/custom_template.py
```

## 🛠️ Troubleshooting

### Common Issues

1. **Email sending fails**
   - Check SMTP settings in config.json
   - Verify email credentials
   - Test with `python main.py test-email`

2. **Certificate generation fails**
   - Check Excel file format and column names
   - Verify image files exist (logo, signature)
   - Check logs in `logs/` directory

3. **Data validation errors**
   - Ensure required columns exist
   - Check for empty values in required fields
   - Verify email address format

### Logging

Check log files in the `logs/` directory for detailed error information:

```bash
tail -f logs/certificate_generator.log
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

- Built with [FPDF](https://pyfpdf.readthedocs.io/) for PDF generation
- Uses [pandas](https://pandas.pydata.org/) for data processing
- Inspired by the need for automated certificate generation in academic research

## 📞 Support

For questions, issues, or contributions:
- Create an issue on GitHub
- Check the examples in the `examples/` directory
- Review the configuration options in `config_example.json`

---

**Happy Certificate Generating! 🎓**
