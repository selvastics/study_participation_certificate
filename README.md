# Certificate System

If you conduct surveys where participants need compensation, such as certificates for study credits in Germany ("Versuchspersonenstunden"), you must keep survey data and personal information separate. All you need to do is store the sensitive data in a separate file and run the proposed solution. It will generate individualized PDF certificates from Excel data and send them automatically. You can also use an LLM to adapt the solution to your specific requirements. If you have any issues, contact me for support.

Generate and send PDF certificates from Excel data.

## Quick Start

```bash
# Install dependencies
pip install fpdf pandas openpyxl

# Create sample data
python create_sample_data.py

# Create example PDFs
python certificate_system.py examples

# Generate certificates
python certificate_system.py generate

# Send emails
python certificate_system.py send

# Run complete workflow
python certificate_system.py workflow
```

## Commands

- `generate` - Generate certificates
- `send` - Send via email
- `workflow` - Complete process
- `examples` - Create example PDFs
- `test-email` - Test email config

## Templates

**Certificate Designs:**
- `default` - Professional standard
- `academic` - Formal university style
- `modern` - Colorful modern design
- `minimal` - Clean simple design

**Email Templates:**
- `default` - English
- `german` - German

## Configuration

Edit `config.json` with your settings:
- Organization name
- Study title
- Instructor name
- Email credentials

## Excel Format

Required columns: `first_name`, `last_name`, `email`

## Files

- `certificate_system.py` - Main system
- `config.json` - Settings
- `create_sample_data.py` - Test data
- `certificates/` - Generated PDFs
- `examples/` - Template examples