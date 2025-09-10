"""
Configuration module for Certificate Generator
Provides centralized configuration management for all components
"""

import os
import json
from typing import Dict, Any, Optional
from dataclasses import dataclass, asdict
from pathlib import Path

@dataclass
class EmailConfig:
    """Email configuration settings"""
    smtp_server: str = ""
    smtp_port: int = 587
    username: str = ""
    password: str = ""
    use_tls: bool = True
    use_ssl: bool = False
    sender_name: str = ""
    sender_email: str = ""

@dataclass
class CertificateConfig:
    """Certificate generation configuration"""
    title: str = "Certificate of Participation"
    subtitle: str = "Study Participation Certificate"
    organization: str = "University"
    department: str = "Psychology"
    instructor_name: str = "Dr. Instructor"
    study_title: str = "Research Study"
    work_unit: str = "Psychology Department"
    hours: float = 1.0
    logo_path: str = "uni.jpg"
    signature_path: str = "sig.jpg"
    output_folder: str = "certificates"
    sent_folder: str = "sended"

@dataclass
class DataConfig:
    """Data processing configuration"""
    excel_file: str = "vpdata.xlsx"
    name_first_column: str = "BE05_01"
    name_last_column: str = "BE05_02"
    email_column: str = "BE05_04"
    id_column: Optional[str] = None

@dataclass
class AppConfig:
    """Main application configuration"""
    email: EmailConfig
    certificate: CertificateConfig
    data: DataConfig
    log_level: str = "INFO"
    delay_between_emails: float = 2.0
    max_retries: int = 3

class ConfigManager:
    """Manages application configuration"""
    
    def __init__(self, config_file: str = "config.json"):
        self.config_file = Path(config_file)
        self.config = self._load_default_config()
        self._load_config()
    
    def _load_default_config(self) -> AppConfig:
        """Load default configuration"""
        return AppConfig(
            email=EmailConfig(),
            certificate=CertificateConfig(),
            data=DataConfig()
        )
    
    def _load_config(self):
        """Load configuration from file or create default"""
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
                self._update_config_from_dict(config_data)
            except Exception as e:
                print(f"Warning: Could not load config file: {e}")
                print("Using default configuration")
        else:
            self.save_config()
    
    def _update_config_from_dict(self, config_data: Dict[str, Any]):
        """Update configuration from dictionary"""
        if 'email' in config_data:
            for key, value in config_data['email'].items():
                if hasattr(self.config.email, key):
                    setattr(self.config.email, key, value)
        
        if 'certificate' in config_data:
            for key, value in config_data['certificate'].items():
                if hasattr(self.config.certificate, key):
                    setattr(self.config.certificate, key, value)
        
        if 'data' in config_data:
            for key, value in config_data['data'].items():
                if hasattr(self.config.data, key):
                    setattr(self.config.data, key, value)
        
        if 'log_level' in config_data:
            self.config.log_level = config_data['log_level']
        if 'delay_between_emails' in config_data:
            self.config.delay_between_emails = config_data['delay_between_emails']
        if 'max_retries' in config_data:
            self.config.max_retries = config_data['max_retries']
    
    def save_config(self):
        """Save current configuration to file"""
        config_dict = {
            'email': asdict(self.config.email),
            'certificate': asdict(self.config.certificate),
            'data': asdict(self.config.data),
            'log_level': self.config.log_level,
            'delay_between_emails': self.config.delay_between_emails,
            'max_retries': self.config.max_retries
        }
        
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(config_dict, f, indent=2, ensure_ascii=False)
    
    def get_config(self) -> AppConfig:
        """Get current configuration"""
        return self.config
    
    def update_config(self, **kwargs):
        """Update configuration values"""
        for key, value in kwargs.items():
            if hasattr(self.config, key):
                setattr(self.config, key, value)
            else:
                print(f"Warning: Unknown configuration key: {key}")
        self.save_config()

# Global config instance
config_manager = ConfigManager()