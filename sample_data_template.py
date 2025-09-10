#!/usr/bin/env python3
"""
Sample data template generator for Certificate Generator
Creates example Excel files with different data structures
"""

import pandas as pd
import os
from pathlib import Path

def create_sample_data():
    """Create sample Excel files for testing"""
    
    # Sample data for basic usage
    basic_data = {
        'BE05_01': ['John', 'Jane', 'Michael', 'Sarah', 'David'],
        'BE05_02': ['Smith', 'Johnson', 'Williams', 'Brown', 'Davis'],
        'BE05_04': ['john.smith@email.com', 'jane.johnson@email.com', 
                    'michael.williams@email.com', 'sarah.brown@email.com', 
                    'david.davis@email.com']
    }
    
    # Sample data with additional columns
    extended_data = {
        'first_name': ['Alice', 'Bob', 'Carol', 'Daniel', 'Eva'],
        'last_name': ['Anderson', 'Baker', 'Clark', 'Davis', 'Evans'],
        'email_address': ['alice.anderson@university.edu', 'bob.baker@university.edu',
                         'carol.clark@university.edu', 'daniel.davis@university.edu',
                         'eva.evans@university.edu'],
        'participant_id': ['P001', 'P002', 'P003', 'P004', 'P005'],
        'study_group': ['Control', 'Treatment', 'Control', 'Treatment', 'Control'],
        'completion_date': ['2024-01-15', '2024-01-16', '2024-01-17', '2024-01-18', '2024-01-19']
    }
    
    # Sample data for German study
    german_data = {
        'BE05_01': ['Anna', 'Max', 'Lisa', 'Tom', 'Sophie'],
        'BE05_02': ['Müller', 'Schmidt', 'Schneider', 'Fischer', 'Weber'],
        'BE05_04': ['anna.mueller@uni-muenster.de', 'max.schmidt@uni-muenster.de',
                   'lisa.schneider@uni-muenster.de', 'tom.fischer@uni-muenster.de',
                   'sophie.weber@uni-muenster.de'],
        'BE05_03': ['P001', 'P002', 'P003', 'P004', 'P005']  # Participant ID
    }
    
    # Create output directory
    output_dir = Path("sample_data")
    output_dir.mkdir(exist_ok=True)
    
    # Save basic data (original format)
    basic_df = pd.DataFrame(basic_data)
    basic_df.to_excel(output_dir / "basic_participants.xlsx", index=False)
    print(f"Created: {output_dir / 'basic_participants.xlsx'}")
    
    # Save extended data
    extended_df = pd.DataFrame(extended_data)
    extended_df.to_excel(output_dir / "extended_participants.xlsx", index=False)
    print(f"Created: {output_dir / 'extended_participants.xlsx'}")
    
    # Save German data
    german_df = pd.DataFrame(german_data)
    german_df.to_excel(output_dir / "german_participants.xlsx", index=False)
    print(f"Created: {output_dir / 'german_participants.xlsx'}")
    
    # Create configuration files for each dataset
    create_config_for_basic(output_dir)
    create_config_for_extended(output_dir)
    create_config_for_german(output_dir)
    
    print(f"\nSample data files created in: {output_dir}")
    print("You can use these files to test the certificate generator with different data formats.")

def create_config_for_basic(output_dir):
    """Create config for basic data format"""
    config = {
        "email": {
            "smtp_server": "smtp.gmail.com",
            "smtp_port": 587,
            "username": "your-email@gmail.com",
            "password": "your-app-password",
            "use_tls": True,
            "use_ssl": False,
            "sender_name": "Research Team",
            "sender_email": "your-email@gmail.com"
        },
        "certificate": {
            "title": "Certificate of Participation",
            "subtitle": "Research Study Participation",
            "organization": "University of Example",
            "department": "Psychology",
            "instructor_name": "Dr. Jane Smith",
            "study_title": "Cognitive Performance Study",
            "work_unit": "Psychology Department",
            "hours": 1.0,
            "logo_path": "uni.jpg",
            "signature_path": "sig.jpg",
            "output_folder": "certificates",
            "sent_folder": "sended"
        },
        "data": {
            "excel_file": "sample_data/basic_participants.xlsx",
            "name_first_column": "BE05_01",
            "name_last_column": "BE05_02",
            "email_column": "BE05_04",
            "id_column": None
        },
        "log_level": "INFO",
        "delay_between_emails": 2.0,
        "max_retries": 3
    }
    
    import json
    with open(output_dir / "config_basic.json", "w") as f:
        json.dump(config, f, indent=2)
    print(f"Created: {output_dir / 'config_basic.json'}")

def create_config_for_extended(output_dir):
    """Create config for extended data format"""
    config = {
        "email": {
            "smtp_server": "smtp.gmail.com",
            "smtp_port": 587,
            "username": "your-email@gmail.com",
            "password": "your-app-password",
            "use_tls": True,
            "use_ssl": False,
            "sender_name": "Research Team",
            "sender_email": "your-email@gmail.com"
        },
        "certificate": {
            "title": "Certificate of Participation",
            "subtitle": "Research Study Participation",
            "organization": "University of Example",
            "department": "Psychology",
            "instructor_name": "Dr. Jane Smith",
            "study_title": "Advanced Cognitive Study",
            "work_unit": "Psychology Department",
            "hours": 2.0,
            "logo_path": "uni.jpg",
            "signature_path": "sig.jpg",
            "output_folder": "certificates",
            "sent_folder": "sended"
        },
        "data": {
            "excel_file": "sample_data/extended_participants.xlsx",
            "name_first_column": "first_name",
            "name_last_column": "last_name",
            "email_column": "email_address",
            "id_column": "participant_id"
        },
        "log_level": "INFO",
        "delay_between_emails": 2.0,
        "max_retries": 3
    }
    
    import json
    with open(output_dir / "config_extended.json", "w") as f:
        json.dump(config, f, indent=2)
    print(f"Created: {output_dir / 'config_extended.json'}")

def create_config_for_german(output_dir):
    """Create config for German data format"""
    config = {
        "email": {
            "smtp_server": "smtp.gmail.com",
            "smtp_port": 587,
            "username": "your-email@gmail.com",
            "password": "your-app-password",
            "use_tls": True,
            "use_ssl": False,
            "sender_name": "Forschungsteam",
            "sender_email": "your-email@gmail.com"
        },
        "certificate": {
            "title": "Nachweis über geleistet Versuchspersonstunden",
            "subtitle": "Studienteilnahme-Zertifikat",
            "organization": "Westfälische Wilhelms-Universität Münster",
            "department": "Psychologie",
            "instructor_name": "Dr. Max Mustermann",
            "study_title": "Kognitionsstudie zur Aufmerksamkeit",
            "work_unit": "Institut für Psychologie",
            "hours": 1.0,
            "logo_path": "uni.jpg",
            "signature_path": "sig.jpg",
            "output_folder": "certificates",
            "sent_folder": "sended"
        },
        "data": {
            "excel_file": "sample_data/german_participants.xlsx",
            "name_first_column": "BE05_01",
            "name_last_column": "BE05_02",
            "email_column": "BE05_04",
            "id_column": "BE05_03"
        },
        "log_level": "INFO",
        "delay_between_emails": 2.0,
        "max_retries": 3
    }
    
    import json
    with open(output_dir / "config_german.json", "w") as f:
        json.dump(config, f, indent=2)
    print(f"Created: {output_dir / 'config_german.json'}")

if __name__ == "__main__":
    create_sample_data()