# Certificate System

Generate and send PDF certificates from Excel data.

## Quick Start

1. **Install dependencies:**
```bash
pip install fpdf pandas openpyxl
```

2. **Create sample data:**
```bash
python create_sample_data.py
```

3. **Configure settings in config.json:**
```json
{
  "email": {
    "username": "your-email@gmail.com",
    "password": "your-app-password"
  },
  "certificate": {
    "title": "Certificate of Participation",
    "organization": "Your University",
    "instructor_name": "Dr. Your Name",
    "study_title": "Your Study Title"
  }
}
```

4. **Run the system:**
```bash
# Generate certificates
python certificate_system.py generate

# Send emails
python certificate_system.py send

# Run complete workflow
python certificate_system.py workflow

# Create example PDFs
python certificate_system.py examples
```

## Commands

- `generate` - Generate certificates from Excel data
- `send` - Send certificates via email
- `organize` - Move sent certificates to sent folder
- `workflow` - Run complete process
- `examples` - Create example PDFs for all templates
- `test-email` - Test email configuration

## Certificate Templates

- `default` - Standard certificate design
- `academic` - Formal academic style
- `modern` - Modern design with colors
- `minimal` - Simple minimal design

## Email Templates

- `default` - English email template
- `german` - German email template

## Excel Format

Required columns:
- first_name
- last_name
- email

## Files

- `certificate_system.py` - Main system (all functionality)
- `config.json` - Configuration settings
- `create_sample_data.py` - Create test data
- `certificates/` - Generated certificates
- `sent/` - Sent certificates
- `examples/` - Example PDFs for each template