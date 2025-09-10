# Quick Start Guide

## 1. Install Dependencies

```bash
pip install fpdf pandas openpyxl
```

## 2. Create Sample Data

```bash
python sample_data_template.py
```

## 3. Configure Settings

```bash
cp config_example.json config.json
```

Edit `config.json` with your email and study details.

## 4. Test the System

```bash
# Test email configuration
python main.py test-email

# Generate certificates from sample data
python main.py generate --excel sample_data/basic_participants.xlsx

# Check status
python main.py status
```

## 5. Run Complete Workflow

```bash
python main.py workflow --excel sample_data/basic_participants.xlsx
```

## Troubleshooting

- Check logs in `logs/certificate_generator.log`
- Verify Excel file has correct column names
- Test email configuration before sending
- Ensure image files (logo, signature) exist