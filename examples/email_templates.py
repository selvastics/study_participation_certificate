#!/usr/bin/env python3
"""
Email templates example for Certificate Generator
Demonstrates how to create and use custom email templates
"""

import sys
import os
from pathlib import Path

# Add parent directory to path to import modules
sys.path.append(str(Path(__file__).parent.parent))

from logger import setup_logger
from email_sender import EmailSender, EmailTemplate
from data_processor import DataProcessor

def create_custom_email_templates():
    """Create custom email templates"""
    
    # Professional template
    professional_template = EmailTemplate(
        subject="Certificate of Completion - {study_title}",
        body="""
        <html>
        <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
            <div style="max-width: 600px; margin: 0 auto; padding: 20px;">
                <h2 style="color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px;">
                    Certificate of Completion
                </h2>
                
                <p>Dear {participant_name},</p>
                
                <p>Congratulations! We are pleased to inform you that you have successfully 
                completed your participation in our research study: <strong>{study_title}</strong>.</p>
                
                <p>Your contribution to our research is invaluable, and we sincerely appreciate 
                the time and effort you dedicated to this study. Please find your official 
                certificate of participation attached to this email.</p>
                
                <div style="background-color: #f8f9fa; padding: 15px; border-left: 4px solid #3498db; margin: 20px 0;">
                    <p style="margin: 0;"><strong>Study Details:</strong></p>
                    <ul style="margin: 10px 0;">
                        <li>Study Title: {study_title}</li>
                        <li>Instructor: {instructor_name}</li>
                        <li>Department: {organization}</li>
                    </ul>
                </div>
                
                <p>If you have any questions about this study or need additional documentation, 
                please don't hesitate to contact us.</p>
                
                <p>Thank you again for your participation!</p>
                
                <p>Best regards,<br>
                <strong>{instructor_name}</strong><br>
                {organization}</p>
                
                <hr style="border: none; border-top: 1px solid #eee; margin: 30px 0;">
                <p style="font-size: 12px; color: #666;">
                    This is an automated message. Please do not reply to this email.
                </p>
            </div>
        </body>
        </html>
        """,
        is_html=True
    )
    
    # Casual template
    casual_template = EmailTemplate(
        subject="Thanks for participating! Your certificate is ready 🎉",
        body="""
        <html>
        <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
            <div style="max-width: 600px; margin: 0 auto; padding: 20px;">
                <h2 style="color: #e74c3c;">Hey {participant_name}! 👋</h2>
                
                <p>Thanks so much for taking part in our study <strong>"{study_title}"</strong>! 
                We really appreciate your time and input.</p>
                
                <p>Your certificate is ready and attached to this email. You can use it for 
                your academic records or just keep it as a memento of your contribution to research! 📜</p>
                
                <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                           color: white; padding: 20px; border-radius: 10px; text-align: center; margin: 20px 0;">
                    <h3 style="margin: 0;">Study: {study_title}</h3>
                    <p style="margin: 5px 0;">Led by: {instructor_name}</p>
                </div>
                
                <p>If you have any questions, just drop us a line!</p>
                
                <p>Cheers! 🎊<br>
                The Research Team</p>
            </div>
        </body>
        </html>
        """,
        is_html=True
    )
    
    # Formal academic template
    formal_template = EmailTemplate(
        subject="Official Certificate of Research Participation - {study_title}",
        body="""
        <html>
        <body style="font-family: 'Times New Roman', serif; line-height: 1.8; color: #000;">
            <div style="max-width: 700px; margin: 0 auto; padding: 40px;">
                <div style="text-align: center; margin-bottom: 40px;">
                    <h1 style="font-size: 24px; margin: 0; color: #1a1a1a;">
                        {organization}
                    </h1>
                    <h2 style="font-size: 18px; margin: 10px 0; color: #666;">
                        Department of {organization}
                    </h2>
                </div>
                
                <p style="text-align: justify; font-size: 14px;">
                    <strong>To Whom It May Concern:</strong>
                </p>
                
                <p style="text-align: justify; font-size: 14px; text-indent: 30px;">
                    This letter serves as official documentation that <strong>{participant_name}</strong> 
                    has successfully completed participation in the research study entitled 
                    <em>"{study_title}"</em> conducted under the supervision of {instructor_name}.
                </p>
                
                <p style="text-align: justify; font-size: 14px; text-indent: 30px;">
                    The participant's contribution to this research was conducted in accordance 
                    with all applicable ethical guidelines and institutional review board 
                    requirements. The attached certificate provides formal recognition of 
                    this participation.
                </p>
                
                <p style="text-align: justify; font-size: 14px; text-indent: 30px;">
                    Should you require any additional information regarding this participation 
                    or the research study, please do not hesitate to contact our office.
                </p>
                
                <div style="margin-top: 60px;">
                    <p style="font-size: 14px;">
                        Sincerely,<br><br>
                        {instructor_name}<br>
                        Principal Investigator<br>
                        {organization}
                    </p>
                </div>
            </div>
        </body>
        </html>
        """,
        is_html=True
    )
    
    return {
        'professional': professional_template,
        'casual': casual_template,
        'formal': formal_template
    }

def main():
    """Email templates example"""
    logger = setup_logger("email_templates_example")
    logger.info("Starting email templates example")
    
    try:
        # Create email sender
        email_sender = EmailSender()
        
        # Add custom templates
        custom_templates = create_custom_email_templates()
        for name, template in custom_templates.items():
            email_sender.add_template(name, template)
            logger.info(f"Added custom template: {name}")
        
        # Show available templates
        available_templates = email_sender.get_available_templates()
        logger.info(f"Available email templates: {', '.join(available_templates)}")
        
        # Load sample data
        data_processor = DataProcessor()
        if not data_processor.load_data("vpdata.xlsx"):
            logger.error("Failed to load data")
            return False
        
        if not data_processor.process_data():
            logger.error("Failed to process data")
            return False
        
        participants = data_processor.get_participants()
        logger.info(f"Loaded {len(participants)} participants")
        
        # Test email configuration
        if not email_sender.test_email_configuration():
            logger.warning("Email configuration test failed - cannot send actual emails")
            logger.info("This example demonstrates template creation and configuration")
            return True
        
        # Send test emails with different templates
        test_participant = participants[0] if participants else {
            'full_name': 'Test Participant',
            'email': 'test@example.com'
        }
        
        # Note: In a real scenario, you would have actual certificate files
        # For this example, we'll just show the template usage
        logger.info("Email templates example completed successfully")
        logger.info("Templates are ready for use with actual certificate files")
        
        return True
        
    except Exception as e:
        logger.error(f"Error in email templates example: {e}")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)