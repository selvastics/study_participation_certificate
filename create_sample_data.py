#!/usr/bin/env python3
"""
Create sample data for testing
"""

import pandas as pd
import os

def create_sample_data():
    """Create sample Excel file"""
    
    # Sample data
    data = {
        'first_name': ['John', 'Jane', 'Michael', 'Sarah', 'David'],
        'last_name': ['Smith', 'Johnson', 'Williams', 'Brown', 'Davis'],
        'email': ['john.smith@email.com', 'jane.johnson@email.com', 
                 'michael.williams@email.com', 'sarah.brown@email.com', 
                 'david.davis@email.com']
    }
    
    # Create DataFrame and save
    df = pd.DataFrame(data)
    df.to_excel('participants.xlsx', index=False)
    print("Created participants.xlsx with sample data")
    
    # Create German data
    german_data = {
        'BE05_01': ['Anna', 'Max', 'Lisa', 'Tom', 'Sophie'],
        'BE05_02': ['Müller', 'Schmidt', 'Schneider', 'Fischer', 'Weber'],
        'BE05_04': ['anna.mueller@uni-muenster.de', 'max.schmidt@uni-muenster.de',
                   'lisa.schneider@uni-muenster.de', 'tom.fischer@uni-muenster.de',
                   'sophie.weber@uni-muenster.de']
    }
    
    df_german = pd.DataFrame(german_data)
    df_german.to_excel('participants_german.xlsx', index=False)
    print("Created participants_german.xlsx with German data")

if __name__ == "__main__":
    create_sample_data()