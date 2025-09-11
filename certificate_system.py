#!/usr/bin/env python3
"""
Certificate System - Consolidated certificate generation and email system
All-in-one solution for generating and sending certificates
"""

import os
import json
import logging
import smtplib
import pandas as pd
from fpdf import FPDF
from datetime import datetime
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email import encoders
import time

class CertificateSystem:
    """Main certificate system class - handles everything"""
    
    def __init__(self, config_file="config.json"):
        self.config_file = config_file
        self.config = self._load_config()
        self._setup_logging()
        self._create_folders()
    
    def _load_config(self):
        """Load configuration from JSON file"""
        default_config = {
            "email": {
                "smtp_server": "smtp.gmail.com",
                "smtp_port": 587,
                "username": "",
                "password": "",
                "sender_name": "Research Team"
            },
            "certificate": {
                "title": "Certificate of Participation",
                "organization": "University",
                "instructor_name": "Dr. Instructor",
                "study_title": "Research Study",
                "hours": 1.0,
                "logo_path": "uni.jpg",
                "signature_path": "sig.jpg"
            },
            "data": {
                "excel_file": "participants.xlsx",
                "name_first_column": "first_name",
                "name_last_column": "last_name",
                "email_column": "email"
            },
            "templates": {
                "certificate": "default",
                "email": "default"
            }
        }
        
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    user_config = json.load(f)
                # Merge with defaults
                for key, value in user_config.items():
                    if key in default_config:
                        default_config[key].update(value)
            except Exception as e:
                print(f"Warning: Could not load config: {e}")
        
        return default_config
    
    def _setup_logging(self):
        """Setup simple logging"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('certificate.log'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def _create_folders(self):
        """Create necessary folders"""
        Path("certificates").mkdir(exist_ok=True)
        Path("sent").mkdir(exist_ok=True)
        Path("examples").mkdir(exist_ok=True)
    
    def load_data(self, excel_file=None):
        """Load and validate participant data"""
        if excel_file is None:
            excel_file = self.config["data"]["excel_file"]
        
        if not os.path.exists(excel_file):
            self.logger.error(f"Excel file not found: {excel_file}")
            return False
        
        try:
            df = pd.read_excel(excel_file)
            self.logger.info(f"Loaded {len(df)} rows from {excel_file}")
            
            # Validate required columns
            required = [
                self.config["data"]["name_first_column"],
                self.config["data"]["name_last_column"],
                self.config["data"]["email_column"]
            ]
            
            missing = [col for col in required if col not in df.columns]
            if missing:
                self.logger.error(f"Missing columns: {missing}")
                return False
            
            # Process data
            self.participants = []
            for idx, row in df.iterrows():
                first_name = str(row[self.config["data"]["name_first_column"]]).strip()
                last_name = str(row[self.config["data"]["name_last_column"]]).strip()
                email = str(row[self.config["data"]["email_column"]]).strip()
                
                if first_name and last_name and email:
                    self.participants.append({
                        'first_name': first_name,
                        'last_name': last_name,
                        'full_name': f"{first_name} {last_name}",
                        'email': email
                    })
            
            self.logger.info(f"Processed {len(self.participants)} valid participants")
            return True
            
        except Exception as e:
            self.logger.error(f"Error loading data: {e}")
            return False
    
    def generate_certificates(self, template="default"):
        """Generate certificates for all participants"""
        if not hasattr(self, 'participants'):
            self.logger.error("No participant data loaded")
            return False
        
        success_count = 0
        for participant in self.participants:
            if self._generate_single_certificate(participant, template):
                success_count += 1
        
        self.logger.info(f"Generated {success_count}/{len(self.participants)} certificates")
        return success_count == len(self.participants)
    
    def _generate_single_certificate(self, participant, template):
        """Generate a single certificate"""
        try:
            pdf = FPDF()
            pdf.add_page()
            
            if template == "default":
                self._create_default_certificate(pdf, participant)
            elif template == "academic":
                self._create_academic_certificate(pdf, participant)
            elif template == "modern":
                self._create_modern_certificate(pdf, participant)
            elif template == "minimal":
                self._create_minimal_certificate(pdf, participant)
            else:
                self._create_default_certificate(pdf, participant)
            
            # Save certificate
            filename = f"{participant['email']}_certificate.pdf"
            filepath = f"certificates/{filename}"
            pdf.output(filepath)
            
            self.logger.info(f"Generated certificate for {participant['full_name']}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error generating certificate for {participant['email']}: {e}")
            return False
    
    def _create_default_certificate(self, pdf, participant):
        """Default certificate template with perfect alignment"""
        # Logo
        if os.path.exists(self.config["certificate"]["logo_path"]):
            pdf.image(self.config["certificate"]["logo_path"], x=20, y=15, w=60)
        
        # Header - Organization (right aligned)
        pdf.set_y(15)
        pdf.set_font("Arial", size=11)
        pdf.cell(0, 10, txt=self.config["certificate"]["organization"], ln=True, align='R')
        
        # Main title
        pdf.ln(25)
        pdf.set_font("Arial", style='B', size=18)
        pdf.cell(200, 12, txt=self.config["certificate"]["title"], ln=True, align='C')
        
        # Subtitle
        pdf.ln(5)
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 8, txt="Research Study Participation Certificate", ln=True, align='C')
        
        # Decorative line
        pdf.ln(10)
        pdf.set_draw_color(0, 0, 0)
        pdf.line(60, pdf.get_y(), 150, pdf.get_y())
        
        # Certificate text
        pdf.ln(20)
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 8, txt="This certifies that", ln=True, align='C')
        
        # Participant name - handle long names
        pdf.ln(8)
        pdf.set_font("Arial", style='B', size=16)
        name = participant['full_name']
        if len(name) > 30:  # Adjust font size for very long names
            pdf.set_font("Arial", style='B', size=14)
        elif len(name) > 25:
            pdf.set_font("Arial", style='B', size=15)
        pdf.cell(200, 10, txt=name, ln=True, align='C')
        
        pdf.ln(8)
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 8, txt="has successfully completed the study", ln=True, align='C')
        
        # Study title - handle long titles
        pdf.ln(5)
        pdf.set_font("Arial", style='B', size=12)
        study_title = self.config["certificate"]["study_title"]
        if len(study_title) > 50:  # Use multi_cell for very long titles
            pdf.multi_cell(180, 8, txt=f'"{study_title}"', align='C')
        else:
            pdf.cell(200, 8, txt=f'"{study_title}"', ln=True, align='C')
        
        # Details box - perfectly aligned with proper spacing
        pdf.ln(10)
        box_x = 25
        box_y = pdf.get_y()
        box_w = 160
        box_h = 35
        
        # Draw box
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(box_x, box_y, box_w, box_h)
        
        # Content inside box - precisely positioned within the box
        pdf.set_font("Arial", size=9)
        line_height = 6
        text_width = box_w - 20  # Account for 10px padding on each side
        
        # Position cursor inside the box - properly positioned
        pdf.set_xy(box_x + 10, box_y + 6)
        
        # Study title - handle long text properly
        study_text = f"Study: {self.config['certificate']['study_title']}"
        pdf.multi_cell(text_width, line_height, txt=study_text, align='L', border=0)
        
        # Instructor - handle long text properly
        instructor_text = f"Instructor: {self.config['certificate']['instructor_name']}"
        pdf.multi_cell(text_width, line_height, txt=instructor_text, align='L', border=0)
        
        # Hours
        pdf.multi_cell(text_width, line_height, txt=f"Hours Completed: {self.config['certificate']['hours']}", align='L', border=0)
        
        # Date
        pdf.multi_cell(text_width, line_height, txt=f"Date: {datetime.now().strftime('%B %d, %Y')}", align='L', border=0)
        
        # Organization - handle long text properly
        org_text = f"Organization: {self.config['certificate']['organization']}"
        pdf.multi_cell(text_width, line_height, txt=org_text, align='L', border=0)
        
        # Signature area - perfectly aligned on same page
        pdf.ln(8)
        
        # Add signature image if available
        if os.path.exists(self.config["certificate"]["signature_path"]):
            pdf.image(self.config["certificate"]["signature_path"], x=30, y=pdf.get_y(), w=50)
        
        # Date area
        pdf.set_xy(115, pdf.get_y())
        pdf.set_font("Arial", size=10)
        pdf.cell(70, 8, txt=f"Date: {datetime.now().strftime('%B %d, %Y')}", align='C')
    
    def _create_academic_certificate(self, pdf, participant):
        """Academic-style certificate template with perfect alignment"""
        # Double border
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(10, 10, 190, 280)
        pdf.rect(15, 15, 180, 270)
        
        # University logo/header
        if os.path.exists(self.config["certificate"]["logo_path"]):
            pdf.image(self.config["certificate"]["logo_path"], x=20, y=25, w=50)
        
        # University name (right aligned)
        pdf.set_y(25)
        pdf.set_font("Times", style='B', size=12)
        pdf.cell(0, 8, txt=self.config["certificate"]["organization"], ln=True, align='R')
        
        # Main title
        pdf.set_y(60)
        pdf.set_font("Times", style='B', size=22)
        pdf.cell(200, 15, txt="CERTIFICATE OF COMPLETION", ln=True, align='C')
        
        # Decorative line
        pdf.ln(8)
        pdf.set_draw_color(0, 0, 0)
        pdf.line(70, pdf.get_y(), 140, pdf.get_y())
        
        # Certificate text
        pdf.ln(25)
        pdf.set_font("Times", size=14)
        pdf.cell(200, 10, txt="This is to certify that", ln=True, align='C')
        
        # Participant name - handle long names
        pdf.ln(12)
        name = participant['full_name']
        if len(name) > 35:  # Adjust font size for very long names
            pdf.set_font("Times", style='B', size=14)
        elif len(name) > 30:
            pdf.set_font("Times", style='B', size=16)
        else:
            pdf.set_font("Times", style='B', size=18)
        pdf.cell(200, 12, txt=name, ln=True, align='C')
        
        pdf.ln(12)
        pdf.set_font("Times", size=14)
        pdf.cell(200, 10, txt="has successfully completed the research study entitled", ln=True, align='C')
        
        # Study title - handle long titles
        pdf.ln(8)
        pdf.set_font("Times", style='I', size=14)
        study_title = self.config["certificate"]["study_title"]
        if len(study_title) > 60:  # Use multi_cell for very long titles
            pdf.multi_cell(180, 10, txt=f'"{study_title}"', align='C')
        else:
            pdf.cell(200, 10, txt=f'"{study_title}"', ln=True, align='C')
        
        # Details box - perfectly aligned with proper spacing
        pdf.ln(12)
        box_x = 20
        box_y = pdf.get_y()
        box_w = 170
        box_h = 40
        
        # Draw box with border
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(box_x, box_y, box_w, box_h)
        
        # Content inside box - precisely positioned within the box
        pdf.set_font("Times", size=9)
        line_height = 5
        text_width = box_w - 20  # Account for 10px padding on each side
        
        # Position cursor inside the box - properly positioned
        pdf.set_xy(box_x + 10, box_y + 6)
        
        # Study title - handle long text properly
        study_text = f"Study Title: {self.config['certificate']['study_title']}"
        pdf.multi_cell(text_width, line_height, txt=study_text, align='L', border=0)
        
        # Principal Investigator - handle long text properly
        pi_text = f"Principal Investigator: {self.config['certificate']['instructor_name']}"
        pdf.multi_cell(text_width, line_height, txt=pi_text, align='L', border=0)
        
        # Hours
        pdf.multi_cell(text_width, line_height, txt=f"Hours Completed: {self.config['certificate']['hours']}", align='L', border=0)
        
        # Institution - handle long text properly
        inst_text = f"Institution: {self.config['certificate']['organization']}"
        pdf.multi_cell(text_width, line_height, txt=inst_text, align='L', border=0)
        
        # Date
        pdf.multi_cell(text_width, line_height, txt=f"Date of Completion: {datetime.now().strftime('%B %d, %Y')}", align='L', border=0)
        
        # Participant ID
        pdf.multi_cell(text_width, line_height, txt=f"Participant ID: {participant['email']}", align='L', border=0)
        
        # Signature area - perfectly aligned on same page
        pdf.ln(8)
        
        # Add signature image if available
        if os.path.exists(self.config["certificate"]["signature_path"]):
            pdf.image(self.config["certificate"]["signature_path"], x=40, y=pdf.get_y(), w=40)
        
        # Date area
        pdf.set_xy(115, pdf.get_y())
        pdf.set_font("Times", size=11)
        pdf.cell(55, 8, txt=f"Date: {datetime.now().strftime('%B %d, %Y')}", align='C')
    
    def _create_modern_certificate(self, pdf, participant):
        """Modern design certificate template with perfect alignment"""
        # Background color (light blue)
        pdf.set_fill_color(240, 248, 255)
        pdf.rect(0, 0, 210, 297, 'F')
        
        # Header with gradient effect
        pdf.set_fill_color(70, 130, 180)
        pdf.rect(0, 0, 210, 70, 'F')
        
        # Logo in header
        if os.path.exists(self.config["certificate"]["logo_path"]):
            pdf.image(self.config["certificate"]["logo_path"], x=20, y=15, w=40)
        
        # Organization name
        pdf.set_y(20)
        pdf.set_font("Arial", style='B', size=12)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(0, 8, txt=self.config["certificate"]["organization"], ln=True, align='R')
        
        # Main title
        pdf.set_y(35)
        pdf.set_font("Arial", style='B', size=20)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(200, 15, txt="CERTIFICATE OF ACHIEVEMENT", ln=True, align='C')
        
        # Reset text color
        pdf.set_text_color(0, 0, 0)
        
        # Main content area
        pdf.set_y(90)
        pdf.set_font("Arial", size=14)
        pdf.cell(200, 10, txt="This certifies that", ln=True, align='C')
        
        # Participant name - handle long names
        pdf.ln(12)
        name = participant['full_name']
        if len(name) > 30:  # Adjust font size for very long names
            pdf.set_font("Arial", style='B', size=14)
        elif len(name) > 25:
            pdf.set_font("Arial", style='B', size=16)
        else:
            pdf.set_font("Arial", style='B', size=18)
        pdf.set_text_color(70, 130, 180)
        pdf.cell(200, 12, txt=name, ln=True, align='C')
        
        pdf.set_text_color(0, 0, 0)
        pdf.ln(12)
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 8, txt="has successfully completed the study", ln=True, align='C')
        
        # Study title - handle long titles
        pdf.ln(8)
        pdf.set_font("Arial", style='B', size=14)
        pdf.set_text_color(70, 130, 180)
        study_title = self.config["certificate"]["study_title"]
        if len(study_title) > 50:  # Use multi_cell for very long titles
            pdf.multi_cell(180, 10, txt=f'"{study_title}"', align='C')
        else:
            pdf.cell(200, 10, txt=f'"{study_title}"', ln=True, align='C')
        
        # Modern info card - perfectly aligned
        pdf.set_text_color(0, 0, 0)
        pdf.ln(12)
        card_x = 25
        card_y = pdf.get_y()
        card_w = 160
        card_h = 45
        
        # Card background
        pdf.set_fill_color(255, 255, 255)
        pdf.set_draw_color(70, 130, 180)
        pdf.rect(card_x, card_y, card_w, card_h, 'FD')
        
        # Card content - precisely positioned within the card
        pdf.set_font("Arial", style='B', size=10)
        pdf.set_text_color(70, 130, 180)
        
        # Position cursor inside the card for header - properly positioned
        pdf.set_xy(card_x + 10, card_y + 8)
        pdf.cell(card_w - 20, 6, txt="STUDY DETAILS", ln=True, align='C')
        
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", size=8)
        line_height = 5
        text_width = card_w - 20  # Account for 10px padding on each side
        
        # Study title - handle long text properly
        study_text = f"Study: {self.config['certificate']['study_title']}"
        pdf.multi_cell(text_width, line_height, txt=study_text, align='L', border=0)
        
        # Instructor - handle long text properly
        instructor_text = f"Instructor: {self.config['certificate']['instructor_name']}"
        pdf.multi_cell(text_width, line_height, txt=instructor_text, align='L', border=0)
        
        # Hours
        pdf.multi_cell(text_width, line_height, txt=f"Hours Completed: {self.config['certificate']['hours']}", align='L', border=0)
        
        # Date
        pdf.multi_cell(text_width, line_height, txt=f"Date: {datetime.now().strftime('%B %d, %Y')}", align='L', border=0)
        
        # Institution - handle long text properly
        inst_text = f"Institution: {self.config['certificate']['organization']}"
        pdf.multi_cell(text_width, line_height, txt=inst_text, align='L', border=0)
        
        # Modern signature area - perfectly aligned on same page
        pdf.ln(8)
        
        # Add signature image if available
        if os.path.exists(self.config["certificate"]["signature_path"]):
            pdf.image(self.config["certificate"]["signature_path"], x=35, y=pdf.get_y(), w=45)
        
        # Date area
        pdf.set_xy(115, pdf.get_y())
        pdf.set_font("Arial", style='B', size=10)
        pdf.set_text_color(70, 130, 180)
        pdf.cell(75, 8, txt=f"Date: {datetime.now().strftime('%B %d, %Y')}", align='C')
    
    def _create_minimal_certificate(self, pdf, participant):
        """Minimal design certificate template with perfect alignment"""
        # Clean, centered design
        pdf.set_y(60)
        pdf.set_font("Arial", style='B', size=16)
        pdf.cell(200, 10, txt="CERTIFICATE", ln=True, align='C')
        
        # Decorative line
        pdf.ln(8)
        pdf.set_draw_color(0, 0, 0)
        pdf.line(80, pdf.get_y(), 130, pdf.get_y())
        
        pdf.ln(25)
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 8, txt="This is to certify that", ln=True, align='C')
        
        # Participant name - handle long names
        pdf.ln(12)
        name = participant['full_name']
        if len(name) > 35:  # Adjust font size for very long names
            pdf.set_font("Arial", style='B', size=12)
        elif len(name) > 30:
            pdf.set_font("Arial", style='B', size=14)
        else:
            pdf.set_font("Arial", style='B', size=16)
        pdf.cell(200, 10, txt=name, ln=True, align='C')
        
        pdf.ln(12)
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 8, txt="has successfully completed", ln=True, align='C')
        
        # Study title - handle long titles
        pdf.ln(8)
        pdf.set_font("Arial", style='B', size=12)
        study_title = self.config["certificate"]["study_title"]
        if len(study_title) > 50:  # Use multi_cell for very long titles
            pdf.multi_cell(180, 8, txt=f'"{study_title}"', align='C')
        else:
            pdf.cell(200, 8, txt=f'"{study_title}"', ln=True, align='C')
        
        # Minimal info section - perfectly aligned
        pdf.ln(15)
        info_x = 35
        info_y = pdf.get_y()
        info_w = 140
        info_h = 30
        
        # Draw subtle border
        pdf.set_draw_color(200, 200, 200)
        pdf.rect(info_x, info_y, info_w, info_h)
        
        # Content inside info box - precisely positioned within the box
        pdf.set_font("Arial", size=8)
        line_height = 5
        text_width = info_w - 20  # Account for 10px padding on each side
        
        # Position cursor inside the box - properly positioned
        pdf.set_xy(info_x + 10, info_y + 6)
        
        # Study title - handle long text properly
        study_text = f"Study: {self.config['certificate']['study_title']}"
        pdf.multi_cell(text_width, line_height, txt=study_text, align='L', border=0)
        
        # Instructor - handle long text properly
        instructor_text = f"Instructor: {self.config['certificate']['instructor_name']}"
        pdf.multi_cell(text_width, line_height, txt=instructor_text, align='L', border=0)
        
        # Hours
        pdf.multi_cell(text_width, line_height, txt=f"Hours Completed: {self.config['certificate']['hours']}", align='L', border=0)
        
        # Date
        pdf.multi_cell(text_width, line_height, txt=f"Date: {datetime.now().strftime('%B %d, %Y')}", align='L', border=0)
        
        # Simple signature area - perfectly aligned on same page
        pdf.ln(8)
        
        # Add signature image if available
        if os.path.exists(self.config["certificate"]["signature_path"]):
            pdf.image(self.config["certificate"]["signature_path"], x=35, y=pdf.get_y(), w=40)
        
        # Date area
        pdf.set_xy(110, pdf.get_y())
        pdf.set_font("Arial", size=10)
        pdf.cell(60, 8, txt=f"Date: {datetime.now().strftime('%B %d, %Y')}", align='C')
    
    def send_emails(self, template="default"):
        """Send certificates via email"""
        if not hasattr(self, 'participants'):
            self.logger.error("No participant data loaded")
            return False
        
        # Test email configuration first
        if not self._test_email():
            return False
        
        success_count = 0
        for participant in self.participants:
            if self._send_single_email(participant, template):
                success_count += 1
            time.sleep(1)  # Delay between emails
        
        self.logger.info(f"Sent {success_count}/{len(self.participants)} emails")
        return success_count == len(self.participants)
    
    def _test_email(self):
        """Test email configuration"""
        try:
            server = smtplib.SMTP(self.config["email"]["smtp_server"], self.config["email"]["smtp_port"])
            server.starttls()
            server.login(self.config["email"]["username"], self.config["email"]["password"])
            server.quit()
            self.logger.info("Email configuration test passed")
            return True
        except Exception as e:
            self.logger.error(f"Email configuration test failed: {e}")
            return False
    
    def _send_single_email(self, participant, template):
        """Send email to single participant"""
        try:
            msg = MIMEMultipart()
            msg['From'] = self.config["email"]["username"]
            msg['To'] = participant['email']
            msg['Subject'] = f"Certificate - {self.config['certificate']['study_title']}"
            
            # Email body
            if template == "german":
                body = f"""
                Liebe:r {participant['first_name']},
                
                vielen Dank für Ihre Teilnahme an unserer Studie "{self.config['certificate']['study_title']}".
                Anbei erhalten Sie Ihr Zertifikat.
                
                Beste Grüße,
                {self.config['certificate']['instructor_name']}
                """
            else:
                body = f"""
                Dear {participant['first_name']},
                
                Thank you for participating in our study "{self.config['certificate']['study_title']}".
                Please find your certificate attached.
                
                Best regards,
                {self.config['certificate']['instructor_name']}
                """
            
            msg.attach(MIMEText(body, 'plain'))
            
            # Attach certificate
            filename = f"{participant['email']}_certificate.pdf"
            filepath = f"certificates/{filename}"
            if os.path.exists(filepath):
                with open(filepath, "rb") as attachment:
                    part = MIMEApplication(attachment.read(), _subtype="pdf")
                    part.add_header('Content-Disposition', f'attachment; filename= {filename}')
                    msg.attach(part)
            
            # Send email
            server = smtplib.SMTP(self.config["email"]["smtp_server"], self.config["email"]["smtp_port"])
            server.starttls()
            server.login(self.config["email"]["username"], self.config["email"]["password"])
            server.send_message(msg)
            server.quit()
            
            self.logger.info(f"Email sent to {participant['email']}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error sending email to {participant['email']}: {e}")
            return False
    
    def organize_files(self):
        """Move sent certificates to sent folder"""
        if not os.path.exists("certificates"):
            return True
        
        moved_count = 0
        for filename in os.listdir("certificates"):
            if filename.endswith("_certificate.pdf"):
                try:
                    os.rename(f"certificates/{filename}", f"sent/{filename}")
                    moved_count += 1
                except Exception as e:
                    self.logger.error(f"Error moving {filename}: {e}")
        
        self.logger.info(f"Moved {moved_count} files to sent folder")
        return True
    
    def create_examples(self):
        """Create example certificates for all templates"""
        example_participant = {
            'first_name': 'Jane',
            'last_name': 'Doe',
            'full_name': 'Jane Doe',
            'email': 'jane.doe@example.com'
        }
        
        templates = ['default', 'academic', 'modern', 'minimal']
        for template in templates:
            try:
                pdf = FPDF()
                pdf.add_page()
                
                if template == "default":
                    self._create_default_certificate(pdf, example_participant)
                elif template == "academic":
                    self._create_academic_certificate(pdf, example_participant)
                elif template == "modern":
                    self._create_modern_certificate(pdf, example_participant)
                elif template == "minimal":
                    self._create_minimal_certificate(pdf, example_participant)
                
                pdf.output(f"examples/example_{template}.pdf")
                self.logger.info(f"Created example: examples/example_{template}.pdf")
                
            except Exception as e:
                self.logger.error(f"Error creating example {template}: {e}")
    
    def run_workflow(self, excel_file=None, cert_template="default", email_template="default"):
        """Run complete workflow"""
        self.logger.info("Starting certificate workflow")
        
        # Load data
        if not self.load_data(excel_file):
            return False
        
        # Generate certificates
        if not self.generate_certificates(cert_template):
            return False
        
        # Send emails
        if not self.send_emails(email_template):
            self.logger.warning("Email sending failed, but certificates were generated")
        
        # Organize files
        self.organize_files()
        
        self.logger.info("Workflow completed successfully")
        return True

def main():
    """Command line interface"""
    import argparse
    
    parser = argparse.ArgumentParser(description="Certificate System")
    parser.add_argument("command", choices=["generate", "send", "organize", "workflow", "examples", "test-email"])
    parser.add_argument("--excel", help="Excel file path")
    parser.add_argument("--cert-template", default="default", choices=["default", "academic", "modern", "minimal"])
    parser.add_argument("--email-template", default="default", choices=["default", "german"])
    
    args = parser.parse_args()
    
    system = CertificateSystem()
    
    if args.command == "generate":
        if system.load_data(args.excel):
            system.generate_certificates(args.cert_template)
    
    elif args.command == "send":
        if system.load_data(args.excel):
            system.send_emails(args.email_template)
    
    elif args.command == "organize":
        system.organize_files()
    
    elif args.command == "workflow":
        system.run_workflow(args.excel, args.cert_template, args.email_template)
    
    elif args.command == "examples":
        system.create_examples()
    
    elif args.command == "test-email":
        system._test_email()

if __name__ == "__main__":
    main()