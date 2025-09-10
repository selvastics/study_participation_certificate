"""
File management module for Certificate Generator
Handles file operations, organization, and cleanup
"""

import os
import shutil
from pathlib import Path
from typing import List, Dict, Any, Optional
from datetime import datetime
from logger import get_logger
from config import config_manager

logger = get_logger(__name__)

class FileManager:
    """Handles file operations and organization"""
    
    def __init__(self):
        self.config = config_manager.get_config()
        self.certificate_folder = Path(self.config.certificate.output_folder)
        self.sent_folder = Path(self.config.certificate.sent_folder)
        
        # Create folders if they don't exist
        self.certificate_folder.mkdir(exist_ok=True)
        self.sent_folder.mkdir(exist_ok=True)
    
    def get_certificate_files(self) -> List[Path]:
        """Get list of certificate files"""
        if not self.certificate_folder.exists():
            return []
        
        return list(self.certificate_folder.glob("*.pdf"))
    
    def get_sent_files(self) -> List[Path]:
        """Get list of sent certificate files"""
        if not self.sent_folder.exists():
            return []
        
        return list(self.sent_folder.glob("*.pdf"))
    
    def move_certificate_to_sent(self, filename: str) -> bool:
        """
        Move certificate file to sent folder
        
        Args:
            filename: Name of the certificate file
        
        Returns:
            True if successful, False otherwise
        """
        source_path = self.certificate_folder / filename
        destination_path = self.sent_folder / filename
        
        if not source_path.exists():
            logger.error(f"Certificate file not found: {source_path}")
            return False
        
        try:
            shutil.move(str(source_path), str(destination_path))
            logger.info(f"Moved {filename} to sent folder")
            return True
        except Exception as e:
            logger.error(f"Error moving {filename}: {e}")
            return False
    
    def move_certificates_to_sent(self, filenames: List[str]) -> Dict[str, Any]:
        """
        Move multiple certificate files to sent folder
        
        Args:
            filenames: List of certificate filenames
        
        Returns:
            Dictionary with move statistics
        """
        results = {
            'total': len(filenames),
            'successful': 0,
            'failed': 0,
            'failed_files': []
        }
        
        logger.info(f"Moving {len(filenames)} certificate files to sent folder")
        
        for filename in filenames:
            if self.move_certificate_to_sent(filename):
                results['successful'] += 1
            else:
                results['failed'] += 1
                results['failed_files'].append(filename)
        
        logger.info(f"Move operation completed: {results['successful']} successful, {results['failed']} failed")
        return results
    
    def move_all_certificates_to_sent(self) -> Dict[str, Any]:
        """Move all certificate files to sent folder"""
        certificate_files = self.get_certificate_files()
        filenames = [f.name for f in certificate_files]
        return self.move_certificates_to_sent(filenames)
    
    def copy_certificate_to_sent(self, filename: str) -> bool:
        """
        Copy certificate file to sent folder (keep original)
        
        Args:
            filename: Name of the certificate file
        
        Returns:
            True if successful, False otherwise
        """
        source_path = self.certificate_folder / filename
        destination_path = self.sent_folder / filename
        
        if not source_path.exists():
            logger.error(f"Certificate file not found: {source_path}")
            return False
        
        try:
            shutil.copy2(str(source_path), str(destination_path))
            logger.info(f"Copied {filename} to sent folder")
            return True
        except Exception as e:
            logger.error(f"Error copying {filename}: {e}")
            return False
    
    def delete_certificate(self, filename: str) -> bool:
        """
        Delete certificate file
        
        Args:
            filename: Name of the certificate file
        
        Returns:
            True if successful, False otherwise
        """
        file_path = self.certificate_folder / filename
        
        if not file_path.exists():
            logger.warning(f"Certificate file not found: {file_path}")
            return True  # Consider it successful if file doesn't exist
        
        try:
            file_path.unlink()
            logger.info(f"Deleted certificate file: {filename}")
            return True
        except Exception as e:
            logger.error(f"Error deleting {filename}: {e}")
            return False
    
    def delete_all_certificates(self) -> int:
        """Delete all certificate files"""
        certificate_files = self.get_certificate_files()
        deleted_count = 0
        
        for file_path in certificate_files:
            if self.delete_certificate(file_path.name):
                deleted_count += 1
        
        logger.info(f"Deleted {deleted_count} certificate files")
        return deleted_count
    
    def cleanup_old_files(self, days: int = 30) -> int:
        """
        Clean up files older than specified days
        
        Args:
            days: Number of days to keep files
        
        Returns:
            Number of files deleted
        """
        cutoff_time = datetime.now().timestamp() - (days * 24 * 60 * 60)
        deleted_count = 0
        
        # Clean up certificate folder
        for file_path in self.certificate_folder.glob("*.pdf"):
            if file_path.stat().st_mtime < cutoff_time:
                if self.delete_certificate(file_path.name):
                    deleted_count += 1
        
        # Clean up sent folder
        for file_path in self.sent_folder.glob("*.pdf"):
            if file_path.stat().st_mtime < cutoff_time:
                try:
                    file_path.unlink()
                    deleted_count += 1
                    logger.info(f"Deleted old sent file: {file_path.name}")
                except Exception as e:
                    logger.error(f"Error deleting old file {file_path.name}: {e}")
        
        logger.info(f"Cleaned up {deleted_count} old files")
        return deleted_count
    
    def get_file_statistics(self) -> Dict[str, Any]:
        """Get file statistics"""
        certificate_files = self.get_certificate_files()
        sent_files = self.get_sent_files()
        
        # Calculate total sizes
        cert_size = sum(f.stat().st_size for f in certificate_files)
        sent_size = sum(f.stat().st_size for f in sent_files)
        
        return {
            'certificate_files': len(certificate_files),
            'sent_files': len(sent_files),
            'certificate_size_mb': round(cert_size / (1024 * 1024), 2),
            'sent_size_mb': round(sent_size / (1024 * 1024), 2),
            'total_size_mb': round((cert_size + sent_size) / (1024 * 1024), 2)
        }
    
    def create_backup(self, backup_folder: str = None) -> bool:
        """
        Create backup of certificate and sent folders
        
        Args:
            backup_folder: Backup folder path (optional)
        
        Returns:
            True if successful, False otherwise
        """
        if backup_folder is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_folder = f"backup_{timestamp}"
        
        backup_path = Path(backup_folder)
        backup_path.mkdir(exist_ok=True)
        
        try:
            # Backup certificate folder
            if self.certificate_folder.exists() and any(self.certificate_folder.iterdir()):
                cert_backup = backup_path / "certificates"
                shutil.copytree(self.certificate_folder, cert_backup)
                logger.info(f"Backed up certificates to {cert_backup}")
            
            # Backup sent folder
            if self.sent_folder.exists() and any(self.sent_folder.iterdir()):
                sent_backup = backup_path / "sended"
                shutil.copytree(self.sent_folder, sent_backup)
                logger.info(f"Backed up sent files to {sent_backup}")
            
            logger.info(f"Backup created successfully: {backup_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error creating backup: {e}")
            return False
    
    def restore_from_backup(self, backup_folder: str) -> bool:
        """
        Restore from backup
        
        Args:
            backup_folder: Backup folder path
        
        Returns:
            True if successful, False otherwise
        """
        backup_path = Path(backup_folder)
        
        if not backup_path.exists():
            logger.error(f"Backup folder not found: {backup_path}")
            return False
        
        try:
            # Restore certificates
            cert_backup = backup_path / "certificates"
            if cert_backup.exists():
                if self.certificate_folder.exists():
                    shutil.rmtree(self.certificate_folder)
                shutil.copytree(cert_backup, self.certificate_folder)
                logger.info("Restored certificates from backup")
            
            # Restore sent files
            sent_backup = backup_path / "sended"
            if sent_backup.exists():
                if self.sent_folder.exists():
                    shutil.rmtree(self.sent_folder)
                shutil.copytree(sent_backup, self.sent_folder)
                logger.info("Restored sent files from backup")
            
            logger.info(f"Restore completed from {backup_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error restoring from backup: {e}")
            return False

# Convenience function
def create_file_manager() -> FileManager:
    """Create and return a new FileManager instance"""
    return FileManager()