"""
Certificate generation module for Certificate Generator
Handles PDF certificate creation with template support
"""

import os
from fpdf import FPDF
from datetime import datetime
from typing import Dict, Any, Optional, List
from pathlib import Path
from logger import get_logger
from config import config_manager

logger = get_logger(__name__)

class CertificateTemplate:
    """Base certificate template class"""
    
    def __init__(self, config):
        self.config = config
    
    def create_certificate(self, pdf: FPDF, participant: Dict[str, Any]) -> None:
        """Create certificate content - to be overridden by subclasses"""
        raise NotImplementedError("Subclasses must implement create_certificate")

class DefaultCertificateTemplate(CertificateTemplate):
    """Default certificate template"""
    
    def create_certificate(self, pdf: FPDF, participant: Dict[str, Any]) -> None:
        """Create default certificate"""
        # Add logo
        if os.path.exists(self.config.logo_path):
            pdf.image(self.config.logo_path, x=20, y=15, w=60)
        else:
            logger.warning(f"Logo file not found: {self.config.logo_path}")
        
        # Add organization info (right-aligned at the top)
        pdf.set_y(15)
        pdf.set_font("Arial", size=11)
        pdf.cell(0, 10, txt=self.config.organization, ln=True, align='R')
        pdf.cell(0, 5, txt=self.config.department, ln=True, align='R')
        
        # Add space between blocks
        pdf.ln(20)
        
        # Add title
        pdf.set_font("Arial", style='B', size=14)
        pdf.cell(200, 10, txt=self.config.title, ln=True, align='C')
        
        # Add subtitle
        pdf.ln(3)
        pdf.set_font("Arial", style='B', size=12)
        pdf.cell(200, 10, txt=self.config.subtitle, ln=True, align='C')
        
        # Add space between blocks
        pdf.ln(10)
        
        # Add certificate details
        self._add_certificate_details(pdf, participant)
        
        # Add signature
        self._add_signature(pdf)
    
    def _add_certificate_details(self, pdf: FPDF, participant: Dict[str, Any]) -> None:
        """Add certificate details section"""
        pdf.set_font("Arial", 'B', size=10)
        
        label_width = 40
        value_width = 110
        
        # Participant name
        pdf.set_x(pdf.get_x() + 20)
        pdf.cell(label_width, 10, txt="Participant Name:", ln=False, align='L')
        pdf.set_font("Arial", size=10)
        pdf.cell(value_width, 10, txt=participant['full_name'], ln=True, align='L')
        pdf.set_font("Arial", 'B', size=10)
        
        # Hours
        pdf.set_x(pdf.get_x() + 20)
        pdf.cell(label_width, 10, txt="Hours:", ln=False, align='L')
        pdf.set_font("Arial", size=10)
        pdf.cell(value_width, 10, txt=str(self.config.hours), ln=True, align='L')
        pdf.set_font("Arial", 'B', size=10)
        
        # Instructor
        pdf.set_x(pdf.get_x() + 20)
        pdf.cell(label_width, 10, txt="Instructor:", ln=False, align='L')
        pdf.set_font("Arial", size=10)
        pdf.cell(value_width, 10, txt=self.config.instructor_name, ln=True, align='L')
        pdf.set_font("Arial", 'B', size=10)
        
        # Study title
        pdf.set_x(pdf.get_x() + 20)
        pdf.cell(label_width, 10, txt="Study Title:", ln=False, align='L')
        pdf.set_font("Arial", size=10)
        pdf.multi_cell(value_width, 10, txt=self.config.study_title, align='L')
        pdf.set_font("Arial", 'B', size=10)
        
        # Work unit
        pdf.set_x(pdf.get_x() + 20)
        pdf.cell(label_width, 10, txt="Department:", ln=False, align='L')
        pdf.set_font("Arial", size=10)
        pdf.cell(value_width, 10, txt=self.config.work_unit, ln=True, align='L')
        pdf.set_font("Arial", 'B', size=10)
        
        # Date
        pdf.set_x(pdf.get_x() + 20)
        pdf.cell(label_width, 10, txt="Date:", ln=False, align='L')
        pdf.set_font("Arial", size=10)
        pdf.cell(value_width, 10, txt=datetime.now().strftime('%Y-%m-%d'), ln=True, align='L')
        pdf.set_font("Arial", 'B', size=10)
        
        # Signature placeholder
        pdf.set_x(pdf.get_x() + 20)
        pdf.cell(label_width, 10, txt="Signature:", ln=False, align='L')
        pdf.set_font("Arial", size=10)
        pdf.cell(10)  # Space for signature
    
    def _add_signature(self, pdf: FPDF) -> None:
        """Add signature section"""
        if os.path.exists(self.config.signature_path):
            pdf.image(self.config.signature_path, x=pdf.get_x(), y=pdf.get_y(), w=60)
            pdf.ln(10)
        else:
            logger.warning(f"Signature file not found: {self.config.signature_path}")

class StudyParticipationTemplate(DefaultCertificateTemplate):
    """Template for study participation certificates"""
    
    def create_certificate(self, pdf: FPDF, participant: Dict[str, Any]) -> None:
        """Create study participation certificate"""
        # Add logo
        if os.path.exists(self.config.logo_path):
            pdf.image(self.config.logo_path, x=20, y=15, w=60)
        
        # Add organization info
        pdf.set_y(15)
        pdf.set_font("Arial", size=11)
        pdf.cell(0, 10, txt=self.config.organization, ln=True, align='R')
        pdf.cell(0, 5, txt=self.config.department, ln=True, align='R')
        
        pdf.ln(20)
        
        # Add title in German (as in original)
        pdf.set_font("Arial", style='B', size=14)
        pdf.cell(200, 10, txt="Nachweis über geleistet Versuchspersonstunden", ln=True, align='C')
        
        pdf.ln(10)
        
        # Add details in German format
        self._add_german_details(pdf, participant)
        self._add_signature(pdf)
    
    def _add_german_details(self, pdf: FPDF, participant: Dict[str, Any]) -> None:
        """Add details in German format"""
        pdf.set_font("Arial", 'B', size=10)
        
        label_width = 40
        value_width = 110
        
        # Name
        pdf.set_x(pdf.get_x() + 20)
        pdf.cell(label_width, 10, txt="Name Student/in:", ln=False, align='L')
        pdf.set_font("Arial", size=10)
        pdf.cell(value_width, 10, txt=participant['full_name'], ln=True, align='L')
        pdf.set_font("Arial", 'B', size=10)
        
        # VP-Stunden
        pdf.set_x(pdf.get_x() + 20)
        pdf.cell(label_width, 10, txt="VP-Stunden:", ln=False, align='L')
        pdf.set_font("Arial", size=10)
        pdf.cell(value_width, 10, txt=str(self.config.hours), ln=True, align='L')
        pdf.set_font("Arial", 'B', size=10)
        
        # Instructor
        pdf.set_x(pdf.get_x() + 20)
        pdf.cell(label_width, 10, txt="Name Dozent/in:", ln=False, align='L')
        pdf.set_font("Arial", size=10)
        pdf.cell(value_width, 10, txt=self.config.instructor_name, ln=True, align='L')
        pdf.set_font("Arial", 'B', size=10)
        
        # Study title
        pdf.set_x(pdf.get_x() + 20)
        pdf.cell(label_width, 10, txt="Titel der Studie:", ln=False, align='L')
        pdf.set_font("Arial", size=10)
        pdf.multi_cell(value_width, 10, txt=self.config.study_title, align='L')
        pdf.set_font("Arial", 'B', size=10)
        
        # Work unit
        pdf.set_x(pdf.get_x() + 20)
        pdf.cell(label_width, 10, txt="Arbeitseinheit:", ln=False, align='L')
        pdf.set_font("Arial", size=10)
        pdf.cell(value_width, 10, txt=self.config.work_unit, ln=True, align='L')
        pdf.set_font("Arial", 'B', size=10)
        
        # Date
        pdf.set_x(pdf.get_x() + 20)
        pdf.cell(label_width, 10, txt="Datum:", ln=False, align='L')
        pdf.set_font("Arial", size=10)
        pdf.cell(value_width, 10, txt=datetime.now().strftime('%Y-%m-%d'), ln=True, align='L')
        pdf.set_font("Arial", 'B', size=10)
        
        # Signature
        pdf.set_x(pdf.get_x() + 20)
        pdf.cell(label_width, 10, txt="Unterschrift Dozent/in:", ln=False, align='L')
        pdf.set_font("Arial", size=10)
        pdf.cell(10)

class CertificateGenerator:
    """Main certificate generator class"""
    
    def __init__(self, template_type: str = "default"):
        self.config = config_manager.get_config().certificate
        self.template = self._get_template(template_type)
        self.output_folder = Path(self.config.output_folder)
        self.output_folder.mkdir(exist_ok=True)
    
    def _get_template(self, template_type: str) -> CertificateTemplate:
        """Get certificate template by type"""
        templates = {
            "default": DefaultCertificateTemplate,
            "study_participation": StudyParticipationTemplate
        }
        
        template_class = templates.get(template_type, DefaultCertificateTemplate)
        return template_class(self.config)
    
    def generate_certificate(self, participant: Dict[str, Any]) -> bool:
        """
        Generate certificate for a single participant
        
        Args:
            participant: Participant data dictionary
        
        Returns:
            True if successful, False otherwise
        """
        try:
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=14)
            pdf.ln(1)
            
            # Create certificate content using template
            self.template.create_certificate(pdf, participant)
            
            # Generate filename
            filename = f"{participant['email']}_certificate.pdf"
            filepath = self.output_folder / filename
            
            # Save certificate
            pdf.output(str(filepath))
            
            logger.info(f"Generated certificate for {participant['full_name']} ({participant['email']})")
            return True
            
        except Exception as e:
            logger.error(f"Error generating certificate for {participant.get('email', 'unknown')}: {e}")
            return False
    
    def generate_certificates(self, participants: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Generate certificates for multiple participants
        
        Args:
            participants: List of participant data dictionaries
        
        Returns:
            Dictionary with generation statistics
        """
        results = {
            'total': len(participants),
            'successful': 0,
            'failed': 0,
            'failed_participants': []
        }
        
        logger.info(f"Starting certificate generation for {len(participants)} participants")
        
        for participant in participants:
            if self.generate_certificate(participant):
                results['successful'] += 1
            else:
                results['failed'] += 1
                results['failed_participants'].append(participant['email'])
        
        logger.info(f"Certificate generation completed: {results['successful']} successful, {results['failed']} failed")
        return results
    
    def get_generated_files(self) -> List[str]:
        """Get list of generated certificate files"""
        if not self.output_folder.exists():
            return []
        
        return [f.name for f in self.output_folder.glob("*.pdf")]
    
    def cleanup_generated_files(self) -> int:
        """Remove all generated certificate files"""
        if not self.output_folder.exists():
            return 0
        
        count = 0
        for file in self.output_folder.glob("*.pdf"):
            file.unlink()
            count += 1
        
        logger.info(f"Cleaned up {count} certificate files")
        return count

# Convenience function
def create_certificate_generator(template_type: str = "default") -> CertificateGenerator:
    """Create and return a new CertificateGenerator instance"""
    return CertificateGenerator(template_type)