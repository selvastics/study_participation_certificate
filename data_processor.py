"""
Data processing module for Certificate Generator
Handles Excel data reading, validation, and processing
"""

import pandas as pd
import os
from typing import List, Dict, Any, Optional, Tuple
from pathlib import Path
import re
from logger import get_logger
from config import config_manager

logger = get_logger(__name__)

class DataProcessor:
    """Handles data processing and validation"""
    
    def __init__(self):
        self.config = config_manager.get_config()
        self.data: Optional[pd.DataFrame] = None
        self.processed_data: List[Dict[str, Any]] = []
    
    def load_data(self, excel_file: Optional[str] = None) -> bool:
        """
        Load data from Excel file
        
        Args:
            excel_file: Path to Excel file (uses config if not provided)
        
        Returns:
            True if successful, False otherwise
        """
        if excel_file is None:
            excel_file = self.config.data.excel_file
        
        try:
            if not os.path.exists(excel_file):
                logger.error(f"Excel file not found: {excel_file}")
                return False
            
            self.data = pd.read_excel(excel_file)
            logger.info(f"Successfully loaded data from {excel_file}")
            logger.info(f"Loaded {len(self.data)} rows")
            return True
            
        except Exception as e:
            logger.error(f"Error loading Excel file: {e}")
            return False
    
    def validate_data(self) -> Tuple[bool, List[str]]:
        """
        Validate loaded data
        
        Returns:
            Tuple of (is_valid, error_messages)
        """
        if self.data is None:
            return False, ["No data loaded"]
        
        errors = []
        config = self.config.data
        
        # Check required columns
        required_columns = [config.name_first_column, config.name_last_column, config.email_column]
        missing_columns = [col for col in required_columns if col not in self.data.columns]
        
        if missing_columns:
            errors.append(f"Missing required columns: {missing_columns}")
        
        # Check for empty values in required columns
        for col in required_columns:
            if col in self.data.columns:
                empty_count = self.data[col].isna().sum()
                if empty_count > 0:
                    errors.append(f"Column '{col}' has {empty_count} empty values")
        
        # Validate email addresses
        if config.email_column in self.data.columns:
            invalid_emails = self._validate_emails(self.data[config.email_column])
            if invalid_emails:
                errors.append(f"Invalid email addresses found: {invalid_emails}")
        
        # Check for duplicate emails
        if config.email_column in self.data.columns:
            duplicates = self.data[config.email_column].duplicated().sum()
            if duplicates > 0:
                errors.append(f"Found {duplicates} duplicate email addresses")
        
        is_valid = len(errors) == 0
        if not is_valid:
            logger.error(f"Data validation failed: {errors}")
        else:
            logger.info("Data validation passed")
        
        return is_valid, errors
    
    def _validate_emails(self, email_series: pd.Series) -> List[str]:
        """Validate email addresses"""
        email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        invalid_emails = []
        
        for idx, email in email_series.items():
            if pd.isna(email):
                continue
            if not re.match(email_pattern, str(email)):
                invalid_emails.append(f"Row {idx + 2}: {email}")  # +2 for header and 0-based index
        
        return invalid_emails
    
    def process_data(self) -> bool:
        """
        Process and clean data for certificate generation
        
        Returns:
            True if successful, False otherwise
        """
        if self.data is None:
            logger.error("No data loaded")
            return False
        
        try:
            config = self.config.data
            self.processed_data = []
            
            for idx, row in self.data.iterrows():
                # Extract and clean data
                first_name = str(row[config.name_first_column]).strip() if pd.notna(row[config.name_first_column]) else ""
                last_name = str(row[config.name_last_column]).strip() if pd.notna(row[config.name_last_column]) else ""
                email = str(row[config.email_column]).strip() if pd.notna(row[config.email_column]) else ""
                
                # Skip rows with missing essential data
                if not first_name or not last_name or not email:
                    logger.warning(f"Skipping row {idx + 2}: Missing essential data")
                    continue
                
                # Create participant data
                participant = {
                    'id': idx + 1,
                    'first_name': first_name,
                    'last_name': last_name,
                    'full_name': f"{first_name} {last_name}",
                    'email': email,
                    'row_index': idx
                }
                
                # Add optional ID column if available
                if config.id_column and config.id_column in self.data.columns:
                    participant['participant_id'] = str(row[config.id_column]).strip() if pd.notna(row[config.id_column]) else ""
                
                self.processed_data.append(participant)
            
            logger.info(f"Processed {len(self.processed_data)} valid participants")
            return True
            
        except Exception as e:
            logger.error(f"Error processing data: {e}")
            return False
    
    def get_participants(self) -> List[Dict[str, Any]]:
        """Get processed participant data"""
        return self.processed_data
    
    def get_participant_by_email(self, email: str) -> Optional[Dict[str, Any]]:
        """Get participant data by email"""
        for participant in self.processed_data:
            if participant['email'] == email:
                return participant
        return None
    
    def export_processed_data(self, output_file: str = "processed_data.csv") -> bool:
        """
        Export processed data to CSV
        
        Args:
            output_file: Output file path
        
        Returns:
            True if successful, False otherwise
        """
        try:
            if not self.processed_data:
                logger.error("No processed data to export")
                return False
            
            df = pd.DataFrame(self.processed_data)
            df.to_csv(output_file, index=False)
            logger.info(f"Exported processed data to {output_file}")
            return True
            
        except Exception as e:
            logger.error(f"Error exporting data: {e}")
            return False
    
    def get_statistics(self) -> Dict[str, Any]:
        """Get data processing statistics"""
        if self.data is None:
            return {}
        
        config = self.config.data
        stats = {
            'total_rows': len(self.data),
            'processed_participants': len(self.processed_data),
            'skipped_rows': len(self.data) - len(self.processed_data),
            'columns': list(self.data.columns),
            'required_columns': [config.name_first_column, config.name_last_column, config.email_column]
        }
        
        return stats

# Convenience function
def create_data_processor() -> DataProcessor:
    """Create and return a new DataProcessor instance"""
    return DataProcessor()