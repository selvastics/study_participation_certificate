# Usage Guide

## Quick Commands

```bash
# Create sample data
python create_sample_data.py

# Generate certificates with different templates
python certificate_system.py generate --cert-template default
python certificate_system.py generate --cert-template academic
python certificate_system.py generate --cert-template modern
python certificate_system.py generate --cert-template minimal

# Create example PDFs for all templates
python certificate_system.py examples

# Send emails
python certificate_system.py send

# Run complete workflow
python certificate_system.py workflow --cert-template modern --email-template german
```

## Certificate Templates

### Default Template
- Professional standard design
- Clean layout with organization header
- Detailed information box
- Perfect for most studies

### Academic Template
- Formal university style
- Double border design
- Times font for academic feel
- Detailed study information
- Perfect for university research

### Modern Template
- Colorful modern design
- Blue header with white text
- Card-style information section
- Contemporary look
- Perfect for modern organizations

### Minimal Template
- Clean, simple design
- Subtle borders
- Centered layout
- Minimal information
- Perfect for simple certificates

## Email Templates

### Default (English)
- Professional English email
- Standard format
- Clear and concise

### German
- German language email
- Formal academic tone
- Perfect for German universities

## Configuration

Edit `config.json` to customize:
- Organization name
- Study title
- Instructor name
- Email settings
- Logo and signature paths

## File Structure

```
certificate_system.py    # Main system
config.json             # Settings
create_sample_data.py   # Create test data
participants.xlsx       # Your data
certificates/           # Generated PDFs
examples/              # Template examples
sent/                  # Sent certificates
```

## Tips

1. Always test with `python certificate_system.py examples` first
2. Check `certificate.log` for any errors
3. Use different templates for different purposes
4. Customize config.json for your organization
5. Test email with `python certificate_system.py test-email`