"""
EC - AI Cost Estimation System Configuration

This module contains configuration settings for the system.
All sensitive information has been removed for security.
"""

import os
from typing import Dict, Any

class Config:
    """Configuration class for the EC AI Cost Estimation System."""
    
    # System Configuration
    SYSTEM_NAME = "EC - AI Cost Estimation System"
    VERSION = "1.0.0"
    
    # AI Model Configuration
    AI_MODEL = "gemini-2.0-flash"
    VISION_DETAIL = "high"
    TEMPERATURE = 1.0
    
    # Excel Template Configuration
    EXCEL_TEMPLATE = {
        'title': 'Cost Sheet',
        'header_rows': 5,
        'data_start_row': 6,
        'font_family': 'Calibri',
        'title_font_size': 26,
        'header_font_size': 11
    }
    
    # Component Categories
    COMPONENT_TYPES = {
        '100': 'Structure',
        '101': 'Furniture&plant (For rent & buy out)',
        '102': 'Graphic',
        '103': 'Electrical'
    }
    
    # Component Codes
    COMPONENT_CODES = {
        '01': 'Flooring',
        '02': 'Main structure & decoration',
        '00': 'Unspecified'
    }
    
    # Units Configuration
    VALID_UNITS = ['sqm', 'unit', 'm', 'pcs']
    
    # Dimension Configuration
    DIMENSIONS = ['W', 'L', 'H']
    
    # API Configuration (placeholder values)
    API_ENDPOINTS = {
        'windmill': 'https://api.example.com/windmill',
        'storage': 'https://storage.example.com'
    }
    
    # Security Configuration
    SECURITY = {
        'encryption_enabled': True,
        'authentication_required': True,
        'data_validation': True
    }
    
    # File Upload Configuration
    FILE_UPLOAD = {
        'allowed_extensions': ['.JPG', '.JPEG', '.PNG', '.GIF', '.WEBP', '.SVG'],
        'max_file_size': 15,  # MB
        'max_files': 10
    }
    
    @classmethod
    def get_ai_config(cls) -> Dict[str, Any]:
        """Get AI model configuration."""
        return {
            'model': cls.AI_MODEL,
            'vision_detail': cls.VISION_DETAIL,
            'temperature': cls.TEMPERATURE
        }
    
    @classmethod
    def get_excel_config(cls) -> Dict[str, Any]:
        """Get Excel template configuration."""
        return cls.EXCEL_TEMPLATE.copy()
    
    @classmethod
    def get_component_types(cls) -> Dict[str, str]:
        """Get component type mappings."""
        return cls.COMPONENT_TYPES.copy()
    
    @classmethod
    def validate_unit(cls, unit: str) -> bool:
        """Validate if unit is supported."""
        return unit in cls.VALID_UNITS
    
    @classmethod
    def validate_dimension(cls, dimension: str) -> bool:
        """Validate if dimension is supported."""
        return dimension in cls.DIMENSIONS

class DevelopmentConfig(Config):
    """Development environment configuration."""
    
    DEBUG = True
    LOG_LEVEL = "DEBUG"
    
    # Development API endpoints
    API_ENDPOINTS = {
        'windmill': 'http://localhost:8000/windmill',
        'storage': 'http://localhost:9000/storage'
    }

class ProductionConfig(Config):
    """Production environment configuration."""
    
    DEBUG = False
    LOG_LEVEL = "INFO"
    
    # Production settings would be loaded from environment variables
    # This is a placeholder structure
    
    @classmethod
    def load_from_env(cls):
        """Load configuration from environment variables."""
        # In production, sensitive data would be loaded from environment
        # This is a demonstration structure
        pass

# Configuration factory
def get_config(environment: str = 'development') -> Config:
    """Get configuration based on environment."""
    if environment == 'production':
        return ProductionConfig()
    else:
        return DevelopmentConfig() 