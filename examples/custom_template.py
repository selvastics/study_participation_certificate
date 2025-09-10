#!/usr/bin/env python3
"""
Custom template example for Certificate Generator
Demonstrates how to create and use custom certificate templates
"""

import sys
import os
from pathlib import Path

# Add parent directory to path to import modules
sys.path.append(str(Path(__file__).parent.parent))

from logger import setup_logger
from certificate_generator import CertificateTemplate, CertificateGenerator
from data_processor import DataProcessor

class CustomCertificateTemplate(CertificateTemplate):
    """Custom certificate template example"""
    
    def create_certificate(self, pdf, participant):
        """Create custom certificate design"""
        # Add custom header
        pdf.set_font("Arial", style='B', size=16)
        pdf.cell(200, 15, txt="CUSTOM ACHIEVEMENT CERTIFICATE", ln=True, align='C')
        
        # Add decorative line
        pdf.ln(5)
        pdf.set_draw_color(0, 0, 255)  # Blue color
        pdf.line(20, pdf.get_y(), 190, pdf.get_y())
        
        # Add participant name with special formatting
        pdf.ln(15)
        pdf.set_font("Arial", style='B', size=14)
        pdf.cell(200, 10, txt=f"This certifies that", ln=True, align='C')
        
        pdf.ln(5)
        pdf.set_font("Arial", style='B', size=18)
        pdf.set_text_color(0, 0, 255)  # Blue text
        pdf.cell(200, 15, txt=participant['full_name'], ln=True, align='C')
        
        # Reset text color
        pdf.set_text_color(0, 0, 0)
        
        # Add achievement description
        pdf.ln(10)
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, txt="has successfully completed", ln=True, align='C')
        
        pdf.ln(5)
        pdf.set_font("Arial", style='B', size=14)
        pdf.cell(200, 10, txt=self.config.study_title, ln=True, align='C')
        
        # Add details in a box
        pdf.ln(15)
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(30, pdf.get_y(), 150, 60)
        
        pdf.set_x(35)
        pdf.set_y(pdf.get_y() + 5)
        pdf.set_font("Arial", size=10)
        
        # Add details inside the box
        details = [
            f"Participant ID: {participant.get('participant_id', 'N/A')}",
            f"Study Duration: {self.config.hours} hours",
            f"Instructor: {self.config.instructor_name}",
            f"Date: {pdf.get_y()}",  # This would need proper date formatting
            f"Department: {self.config.department}"
        ]
        
        for detail in details:
            pdf.cell(140, 8, txt=detail, ln=True)
        
        # Add signature area
        pdf.ln(20)
        pdf.set_font("Arial", style='B', size=10)
        pdf.cell(60, 10, txt="Instructor Signature:", ln=False)
        pdf.cell(60, 10, txt="Date:", ln=True)
        
        # Add signature line
        pdf.line(20, pdf.get_y() + 5, 80, pdf.get_y() + 5)
        pdf.line(100, pdf.get_y() - 5, 160, pdf.get_y() - 5)

def main():
    """Custom template example"""
    logger = setup_logger("custom_template_example")
    logger.info("Starting custom template example")
    
    try:
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
        
        # Create custom generator with custom template
        cert_generator = CertificateGenerator("default")
        
        # Replace the template with our custom one
        cert_generator.template = CustomCertificateTemplate(cert_generator.config)
        
        # Generate certificates
        results = cert_generator.generate_certificates(participants)
        logger.info(f"Generated {results['successful']} custom certificates")
        
        if results['failed'] > 0:
            logger.warning(f"Failed to generate {results['failed']} certificates")
        
        logger.info("Custom template example completed")
        return True
        
    except Exception as e:
        logger.error(f"Error in custom template example: {e}")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)