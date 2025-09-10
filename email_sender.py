"""
Email sending module for Certificate Generator
Handles email delivery with multiple backend support
"""

import os
import smtplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email import encoders
from typing import Dict, Any, List, Optional, Tuple
from pathlib import Path
from logger import get_logger
from config import config_manager

logger = get_logger(__name__)

class EmailTemplate:
    """Email template class"""
    
    def __init__(self, subject: str, body: str, is_html: bool = True):
        self.subject = subject
        self.body = body
        self.is_html = is_html
    
    def format(self, **kwargs) -> Tuple[str, str]:
        """Format template with provided variables"""
        formatted_subject = self.subject.format(**kwargs)
        formatted_body = self.body.format(**kwargs)
        return formatted_subject, formatted_body

class EmailSender:
    """Main email sending class"""
    
    def __init__(self):
        self.config = config_manager.get_config()
        self.email_config = self.config.email
        self.delay = self.config.delay_between_emails
        self.max_retries = self.config.max_retries
        
        # Default email templates
        self.templates = {
            'default': EmailTemplate(
                subject="Certificate of Participation - {study_title}",
                body="""
                <html>
                <body>
                <p>Dear {participant_name},</p>
                
                <p>Thank you for your participation in our study "{study_title}". 
                Please find attached your certificate of participation.</p>
                
                <p>If you have any questions, please don't hesitate to contact us.</p>
                
                <p>Best regards,<br>
                {instructor_name}<br>
                {organization}</p>
                </body>
                </html>
                """,
                is_html=True
            ),
            'german': EmailTemplate(
                subject="Teilnahme an der {study_title} und geleistet Versuchspersonstunden",
                body="""
                <html>
                <body>
                <p><strong>Liebe:r Teilnehmer:in,</strong></p>
                
                <p>vielen Dank für Ihre Teilnahme an unserer Studie zu {study_title}. 
                Anbei erhalten Sie Ihren Nachweis über die geleistet Versuchspersonstunde. 
                Sollten Sie Fragen haben, stehen wir Ihnen gerne zur Verfügung.</p>
                
                <p>Beste Grüße,<br>
                {instructor_name}</p>
                </body>
                </html>
                """,
                is_html=True
            ),
            'simple': EmailTemplate(
                subject="Your Certificate - {study_title}",
                body="Dear {participant_name},\n\nThank you for participating in our study. Please find your certificate attached.\n\nBest regards,\n{instructor_name}",
                is_html=False
            )
        }
    
    def _create_smtp_connection(self) -> smtplib.SMTP:
        """Create SMTP connection"""
        if self.email_config.use_ssl:
            server = smtplib.SMTP_SSL(self.email_config.smtp_server, self.email_config.smtp_port)
        else:
            server = smtplib.SMTP(self.email_config.smtp_server, self.email_config.smtp_port)
        
        if self.email_config.use_tls and not self.email_config.use_ssl:
            server.starttls()
        
        return server
    
    def _send_email_smtp(self, to_email: str, subject: str, body: str, 
                        attachment_path: Optional[str] = None, is_html: bool = True) -> bool:
        """Send email using SMTP"""
        try:
            # Create message
            msg = MIMEMultipart()
            msg['From'] = self.email_config.sender_email or self.email_config.username
            msg['To'] = to_email
            msg['Subject'] = subject
            
            # Add body
            if is_html:
                msg.attach(MIMEText(body, 'html'))
            else:
                msg.attach(MIMEText(body, 'plain'))
            
            # Add attachment if provided
            if attachment_path and os.path.exists(attachment_path):
                with open(attachment_path, "rb") as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename= {os.path.basename(attachment_path)}'
                    )
                    msg.attach(part)
            
            # Send email
            with self._create_smtp_connection() as server:
                server.login(self.email_config.username, self.email_config.password)
                server.send_message(msg)
            
            return True
            
        except Exception as e:
            logger.error(f"Error sending email to {to_email}: {e}")
            return False
    
    def send_email(self, to_email: str, subject: str, body: str, 
                  attachment_path: Optional[str] = None, is_html: bool = True) -> bool:
        """
        Send email with retry logic
        
        Args:
            to_email: Recipient email address
            subject: Email subject
            body: Email body
            attachment_path: Path to attachment file
            is_html: Whether body is HTML
        
        Returns:
            True if successful, False otherwise
        """
        for attempt in range(self.max_retries):
            try:
                if self._send_email_smtp(to_email, subject, body, attachment_path, is_html):
                    logger.info(f"Email sent successfully to {to_email}")
                    return True
                else:
                    logger.warning(f"Failed to send email to {to_email} (attempt {attempt + 1})")
                    
            except Exception as e:
                logger.error(f"Error sending email to {to_email} (attempt {attempt + 1}): {e}")
            
            if attempt < self.max_retries - 1:
                time.sleep(self.delay)
        
        logger.error(f"Failed to send email to {to_email} after {self.max_retries} attempts")
        return False
    
    def send_certificate_email(self, participant: Dict[str, Any], 
                             certificate_path: str, 
                             template_name: str = "default") -> bool:
        """
        Send certificate email to participant
        
        Args:
            participant: Participant data
            certificate_path: Path to certificate PDF
            template_name: Email template to use
        
        Returns:
            True if successful, False otherwise
        """
        if template_name not in self.templates:
            logger.error(f"Unknown email template: {template_name}")
            return False
        
        template = self.templates[template_name]
        
        # Format template variables
        template_vars = {
            'participant_name': participant['full_name'],
            'study_title': self.config.certificate.study_title,
            'instructor_name': self.config.certificate.instructor_name,
            'organization': self.config.certificate.organization
        }
        
        subject, body = template.format(**template_vars)
        
        return self.send_email(
            to_email=participant['email'],
            subject=subject,
            body=body,
            attachment_path=certificate_path,
            is_html=template.is_html
        )
    
    def send_bulk_certificates(self, participants: List[Dict[str, Any]], 
                             certificate_folder: str,
                             template_name: str = "default") -> Dict[str, Any]:
        """
        Send certificates to multiple participants
        
        Args:
            participants: List of participant data
            certificate_folder: Folder containing certificate files
            template_name: Email template to use
        
        Returns:
            Dictionary with sending statistics
        """
        results = {
            'total': len(participants),
            'successful': 0,
            'failed': 0,
            'failed_participants': []
        }
        
        logger.info(f"Starting bulk email sending for {len(participants)} participants")
        
        for participant in participants:
            certificate_filename = f"{participant['email']}_certificate.pdf"
            certificate_path = os.path.join(certificate_folder, certificate_filename)
            
            if not os.path.exists(certificate_path):
                logger.error(f"Certificate not found: {certificate_path}")
                results['failed'] += 1
                results['failed_participants'].append(participant['email'])
                continue
            
            if self.send_certificate_email(participant, certificate_path, template_name):
                results['successful'] += 1
            else:
                results['failed'] += 1
                results['failed_participants'].append(participant['email'])
            
            # Delay between emails
            if self.delay > 0:
                time.sleep(self.delay)
        
        logger.info(f"Bulk email sending completed: {results['successful']} successful, {results['failed']} failed")
        return results
    
    def test_email_configuration(self) -> bool:
        """
        Test email configuration by sending a test email
        
        Returns:
            True if configuration is valid, False otherwise
        """
        try:
            test_subject = "Certificate Generator - Test Email"
            test_body = "This is a test email to verify your email configuration."
            
            # Send test email to sender's own address
            test_email = self.email_config.sender_email or self.email_config.username
            
            if self.send_email(test_email, test_subject, test_body):
                logger.info("Email configuration test successful")
                return True
            else:
                logger.error("Email configuration test failed")
                return False
                
        except Exception as e:
            logger.error(f"Email configuration test error: {e}")
            return False
    
    def add_template(self, name: str, template: EmailTemplate) -> None:
        """Add custom email template"""
        self.templates[name] = template
        logger.info(f"Added email template: {name}")
    
    def get_available_templates(self) -> List[str]:
        """Get list of available email templates"""
        return list(self.templates.keys())

# Convenience function
def create_email_sender() -> EmailSender:
    """Create and return a new EmailSender instance"""
    return EmailSender()