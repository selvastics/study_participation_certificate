#!/usr/bin/env python3
"""
Basic usage example for Certificate Generator
Demonstrates simple certificate generation and email sending
"""

import sys
import os
from pathlib import Path

# Add parent directory to path to import modules
sys.path.append(str(Path(__file__).parent.parent))

from logger import setup_logger
from data_processor import DataProcessor
from certificate_generator import CertificateGenerator
from email_sender import EmailSender
from file_manager import FileManager

def main():
    """Basic usage example"""
    # Setup logging
    logger = setup_logger("basic_example", log_file="logs/basic_example.log")
    logger.info("Starting basic usage example")
    
    try:
        # Step 1: Process data
        logger.info("Step 1: Processing data")
        data_processor = DataProcessor()
        
        if not data_processor.load_data("vpdata.xlsx"):
            logger.error("Failed to load data")
            return False
        
        if not data_processor.process_data():
            logger.error("Failed to process data")
            return False
        
        participants = data_processor.get_participants()
        logger.info(f"Loaded {len(participants)} participants")
        
        # Step 2: Generate certificates
        logger.info("Step 2: Generating certificates")
        cert_generator = CertificateGenerator("study_participation")
        
        results = cert_generator.generate_certificates(participants)
        logger.info(f"Generated {results['successful']} certificates")
        
        # Step 3: Send emails (optional - requires email configuration)
        logger.info("Step 3: Sending emails")
        email_sender = EmailSender()
        
        # Test email configuration first
        if email_sender.test_email_configuration():
            email_results = email_sender.send_bulk_certificates(
                participants, 
                str(cert_generator.output_folder),
                "german"
            )
            logger.info(f"Sent {email_results['successful']} emails")
        else:
            logger.warning("Email configuration test failed - skipping email sending")
        
        # Step 4: Organize files
        logger.info("Step 4: Organizing files")
        file_manager = FileManager()
        file_manager.move_all_certificates_to_sent()
        
        logger.info("Basic usage example completed successfully")
        return True
        
    except Exception as e:
        logger.error(f"Error in basic example: {e}")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)