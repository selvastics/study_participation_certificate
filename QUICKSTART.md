# Quick Start Guide

Get up and running with Certificate Generator in 5 minutes!

## 1. Install Dependencies

```bash
pip install -r requirements.txt
```

## 2. Create Sample Data

```bash
python sample_data_template.py
```

This creates sample Excel files and configuration files in the `sample_data/` directory.

## 3. Configure Your Settings

```bash
cp config_example.json config.json
```

Edit `config.json` with your email settings and study details.

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
# Run everything at once
python main.py workflow --excel sample_data/basic_participants.xlsx
```

## 6. Customize for Your Needs

- Edit `config.json` for your specific study
- Modify certificate templates in `certificate_generator.py`
- Create custom email templates
- Add your own logo and signature images

## Troubleshooting

- Check logs in `logs/certificate_generator.log`
- Verify your Excel file has the correct column names
- Test email configuration before sending bulk emails
- Ensure image files (logo, signature) exist

## Next Steps

- Read the full [README.md](README.md) for detailed documentation
- Check out examples in the `examples/` directory
- Customize templates for your specific needs

Happy certificate generating! 🎓