#!/usr/bin/env python3
"""
Main orchestration script for Certificate Generator
Provides a unified interface for all certificate generation operations
"""

import argparse
import sys
from pathlib import Path
from typing import List, Dict, Any, Optional

from logger import setup_logger, get_logger
from config import config_manager
from data_processor import DataProcessor
from certificate_generator import CertificateGenerator
from email_sender import EmailSender
from file_manager import FileManager

def setup_application():
    """Initialize application components"""
    logger = setup_logger("certificate_generator", log_file="logs/certificate_generator.log")
    logger.info("Certificate Generator started")
    return logger

def generate_certificates(excel_file: Optional[str] = None, 
                         template_type: str = "default",
                         output_folder: Optional[str] = None) -> bool:
    """
    Generate certificates from Excel data
    
    Args:
        excel_file: Path to Excel file
        template_type: Certificate template type
        output_folder: Output folder for certificates
    
    Returns:
        True if successful, False otherwise
    """
    logger = get_logger()
    
    try:
        # Initialize components
        data_processor = DataProcessor()
        cert_generator = CertificateGenerator(template_type)
        
        # Load and validate data
        if not data_processor.load_data(excel_file):
            return False
        
        is_valid, errors = data_processor.validate_data()
        if not is_valid:
            logger.error("Data validation failed:")
            for error in errors:
                logger.error(f"  - {error}")
            return False
        
        # Process data
        if not data_processor.process_data():
            return False
        
        # Generate certificates
        participants = data_processor.get_participants()
        if not participants:
            logger.error("No valid participants found")
            return False
        
        results = cert_generator.generate_certificates(participants)
        
        logger.info(f"Certificate generation completed:")
        logger.info(f"  - Total participants: {results['total']}")
        logger.info(f"  - Successful: {results['successful']}")
        logger.info(f"  - Failed: {results['failed']}")
        
        if results['failed_participants']:
            logger.warning("Failed participants:")
            for email in results['failed_participants']:
                logger.warning(f"  - {email}")
        
        return results['failed'] == 0
        
    except Exception as e:
        logger.error(f"Error in certificate generation: {e}")
        return False

def send_certificates(template_name: str = "default",
                     test_mode: bool = False) -> bool:
    """
    Send certificates via email
    
    Args:
        template_name: Email template to use
        test_mode: If True, only test email configuration
    
    Returns:
        True if successful, False otherwise
    """
    logger = get_logger()
    
    try:
        # Initialize components
        email_sender = EmailSender()
        file_manager = FileManager()
        
        # Test email configuration
        if not email_sender.test_email_configuration():
            logger.error("Email configuration test failed")
            return False
        
        if test_mode:
            logger.info("Email configuration test passed")
            return True
        
        # Get certificate files
        certificate_files = file_manager.get_certificate_files()
        if not certificate_files:
            logger.warning("No certificate files found to send")
            return True
        
        # Load participant data
        data_processor = DataProcessor()
        if not data_processor.load_data():
            return False
        
        if not data_processor.process_data():
            return False
        
        participants = data_processor.get_participants()
        
        # Send certificates
        results = email_sender.send_bulk_certificates(
            participants, 
            str(file_manager.certificate_folder),
            template_name
        )
        
        logger.info(f"Email sending completed:")
        logger.info(f"  - Total participants: {results['total']}")
        logger.info(f"  - Successful: {results['successful']}")
        logger.info(f"  - Failed: {results['failed']}")
        
        if results['failed_participants']:
            logger.warning("Failed to send emails to:")
            for email in results['failed_participants']:
                logger.warning(f"  - {email}")
        
        return results['failed'] == 0
        
    except Exception as e:
        logger.error(f"Error in email sending: {e}")
        return False

def organize_files(move_to_sent: bool = True) -> bool:
    """
    Organize certificate files
    
    Args:
        move_to_sent: If True, move sent files to sent folder
    
    Returns:
        True if successful, False otherwise
    """
    logger = get_logger()
    
    try:
        file_manager = FileManager()
        
        if move_to_sent:
            results = file_manager.move_all_certificates_to_sent()
            logger.info(f"File organization completed:")
            logger.info(f"  - Total files: {results['total']}")
            logger.info(f"  - Moved successfully: {results['successful']}")
            logger.info(f"  - Failed: {results['failed']}")
            
            if results['failed_files']:
                logger.warning("Failed to move files:")
                for filename in results['failed_files']:
                    logger.warning(f"  - {filename}")
            
            return results['failed'] == 0
        else:
            stats = file_manager.get_file_statistics()
            logger.info("File statistics:")
            logger.info(f"  - Certificate files: {stats['certificate_files']}")
            logger.info(f"  - Sent files: {stats['sent_files']}")
            logger.info(f"  - Total size: {stats['total_size_mb']} MB")
            return True
        
    except Exception as e:
        logger.error(f"Error in file organization: {e}")
        return False

def full_workflow(excel_file: Optional[str] = None,
                 template_type: str = "default",
                 email_template: str = "default",
                 move_files: bool = True) -> bool:
    """
    Run complete certificate generation workflow
    
    Args:
        excel_file: Path to Excel file
        template_type: Certificate template type
        email_template: Email template type
        move_files: Whether to move sent files to sent folder
    
    Returns:
        True if successful, False otherwise
    """
    logger = get_logger()
    logger.info("Starting full certificate generation workflow")
    
    # Step 1: Generate certificates
    logger.info("Step 1: Generating certificates")
    if not generate_certificates(excel_file, template_type):
        logger.error("Certificate generation failed")
        return False
    
    # Step 2: Send emails
    logger.info("Step 2: Sending emails")
    if not send_certificates(email_template):
        logger.error("Email sending failed")
        return False
    
    # Step 3: Organize files
    if move_files:
        logger.info("Step 3: Organizing files")
        if not organize_files(True):
            logger.warning("File organization had issues, but continuing")
    
    logger.info("Full workflow completed successfully")
    return True

def show_status() -> None:
    """Show application status and statistics"""
    logger = get_logger()
    
    try:
        # File statistics
        file_manager = FileManager()
        file_stats = file_manager.get_file_statistics()
        
        # Data statistics
        data_processor = DataProcessor()
        if data_processor.load_data():
            data_processor.process_data()
            data_stats = data_processor.get_statistics()
        else:
            data_stats = {}
        
        logger.info("=== Certificate Generator Status ===")
        logger.info(f"Certificate files: {file_stats['certificate_files']}")
        logger.info(f"Sent files: {file_stats['sent_files']}")
        logger.info(f"Total size: {file_stats['total_size_mb']} MB")
        
        if data_stats:
            logger.info(f"Excel rows: {data_stats['total_rows']}")
            logger.info(f"Valid participants: {data_stats['processed_participants']}")
            logger.info(f"Skipped rows: {data_stats['skipped_rows']}")
        
        # Configuration info
        config = config_manager.get_config()
        logger.info(f"Certificate template: {config.certificate.title}")
        logger.info(f"Study title: {config.certificate.study_title}")
        logger.info(f"Instructor: {config.certificate.instructor_name}")
        
    except Exception as e:
        logger.error(f"Error showing status: {e}")

def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(
        description="Certificate Generator - Automated certificate creation and distribution",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python main.py generate                    # Generate certificates
  python main.py send                        # Send certificates via email
  python main.py organize                    # Organize files
  python main.py workflow                    # Run complete workflow
  python main.py status                      # Show status
  python main.py test-email                  # Test email configuration
        """
    )
    
    subparsers = parser.add_subparsers(dest='command', help='Available commands')
    
    # Generate command
    gen_parser = subparsers.add_parser('generate', help='Generate certificates')
    gen_parser.add_argument('--excel', help='Excel file path')
    gen_parser.add_argument('--template', default='default', 
                           choices=['default', 'study_participation'],
                           help='Certificate template type')
    gen_parser.add_argument('--output', help='Output folder')
    
    # Send command
    send_parser = subparsers.add_parser('send', help='Send certificates via email')
    send_parser.add_argument('--template', default='default',
                            help='Email template name')
    send_parser.add_argument('--test', action='store_true',
                            help='Test email configuration only')
    
    # Organize command
    org_parser = subparsers.add_parser('organize', help='Organize certificate files')
    org_parser.add_argument('--no-move', action='store_true',
                           help='Don\'t move files, just show statistics')
    
    # Workflow command
    workflow_parser = subparsers.add_parser('workflow', help='Run complete workflow')
    workflow_parser.add_argument('--excel', help='Excel file path')
    workflow_parser.add_argument('--cert-template', default='default',
                                choices=['default', 'study_participation'],
                                help='Certificate template type')
    workflow_parser.add_argument('--email-template', default='default',
                                help='Email template name')
    workflow_parser.add_argument('--no-move', action='store_true',
                                help='Don\'t move files after sending')
    
    # Status command
    subparsers.add_parser('status', help='Show application status')
    
    # Test email command
    subparsers.add_parser('test-email', help='Test email configuration')
    
    args = parser.parse_args()
    
    # Setup application
    logger = setup_application()
    
    try:
        if args.command == 'generate':
            success = generate_certificates(
                excel_file=args.excel,
                template_type=args.template,
                output_folder=args.output
            )
            sys.exit(0 if success else 1)
            
        elif args.command == 'send':
            success = send_certificates(
                template_name=args.template,
                test_mode=args.test
            )
            sys.exit(0 if success else 1)
            
        elif args.command == 'organize':
            success = organize_files(move_to_sent=not args.no_move)
            sys.exit(0 if success else 1)
            
        elif args.command == 'workflow':
            success = full_workflow(
                excel_file=args.excel,
                template_type=args.cert_template,
                email_template=args.email_template,
                move_files=not args.no_move
            )
            sys.exit(0 if success else 1)
            
        elif args.command == 'status':
            show_status()
            sys.exit(0)
            
        elif args.command == 'test-email':
            success = send_certificates(test_mode=True)
            sys.exit(0 if success else 1)
            
        else:
            parser.print_help()
            sys.exit(1)
            
    except KeyboardInterrupt:
        logger.info("Operation cancelled by user")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Unexpected error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()