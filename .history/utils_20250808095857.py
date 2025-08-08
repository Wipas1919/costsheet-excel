"""
EC - AI Cost Estimation System Utilities

Helper functions and utilities for the cost estimation system.
"""

import re
import json
from typing import Dict, List, Any, Optional
from datetime import datetime
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def validate_list_id(list_id: str) -> bool:
    """
    Validates the format of list_id.
    
    Args:
        list_id (str): List ID to validate
        
    Returns:
        bool: True if valid, False otherwise
    """
    # Expected format: XXX-XX-XX (e.g., 100-01-01)
    pattern = r'^\d{3}-\d{2}-\d{2}$'
    return bool(re.match(pattern, list_id))

def parse_list_id(list_id: str) -> Dict[str, str]:
    """
    Parses list_id into its components.
    
    Args:
        list_id (str): List ID to parse
        
    Returns:
        dict: Parsed components
    """
    if not validate_list_id(list_id):
        return {}
    
    parts = list_id.split('-')
    return {
        'type_code': parts[0],
        'component_code': parts[1],
        'item_number': parts[2]
    }

def calculate_surface_area(width: float, length: float, height: float = None) -> float:
    """
    Calculates surface area based on dimensions.
    
    Args:
        width (float): Width in meters
        length (float): Length in meters
        height (float, optional): Height in meters
        
    Returns:
        float: Calculated surface area
    """
    if height is None:
        # For flooring, use width × length
        return width * length
    else:
        # For walls/structures, use perimeter × height
        perimeter = 2 * (width + length)
        return perimeter * height

def validate_dimensions(width: str, length: str, height: str = None) -> Dict[str, Any]:
    """
    Validates and converts dimension strings to floats.
    
    Args:
        width (str): Width dimension
        length (str): Length dimension
        height (str, optional): Height dimension
        
    Returns:
        dict: Validation result with converted values
    """
    result = {
        'valid': True,
        'width': None,
        'length': None,
        'height': None,
        'errors': []
    }
    
    try:
        result['width'] = float(width) if width and width != '-' else None
    except (ValueError, TypeError):
        result['errors'].append(f"Invalid width: {width}")
        result['valid'] = False
    
    try:
        result['length'] = float(length) if length and length != '-' else None
    except (ValueError, TypeError):
        result['errors'].append(f"Invalid length: {length}")
        result['valid'] = False
    
    if height:
        try:
            result['height'] = float(height) if height and height != '-' else None
        except (ValueError, TypeError):
            result['errors'].append(f"Invalid height: {height}")
            result['valid'] = False
    
    return result

def calculate_quantity(component_type: str, width: float, length: float, 
                      height: float = None, unit: str = 'sqm') -> float:
    """
    Calculates quantity based on component type and dimensions.
    
    Args:
        component_type (str): Type of component
        width (float): Width in meters
        length (float): Length in meters
        height (float, optional): Height in meters
        unit (str): Unit of measurement
        
    Returns:
        float: Calculated quantity
    """
    if unit == 'sqm':
        if 'Flooring' in component_type:
            return width * length
        elif 'Structure' in component_type and height:
            # For walls/structures, calculate surface area
            return calculate_surface_area(width, length, height)
        else:
            return width * length
    elif unit == 'unit':
        return 1.0
    else:
        return 0.0

def format_currency(amount: float) -> str:
    """
    Formats currency amount with proper formatting.
    
    Args:
        amount (float): Amount to format
        
    Returns:
        str: Formatted currency string
    """
    try:
        return f"{amount:,.2f}"
    except (ValueError, TypeError):
        return "0.00"

def sanitize_text(text: str) -> str:
    """
    Sanitizes text input for safe processing.
    
    Args:
        text (str): Text to sanitize
        
    Returns:
        str: Sanitized text
    """
    if not text:
        return '-'
    
    # Remove potentially dangerous characters
    sanitized = re.sub(r'[<>"\']', '', str(text))
    return sanitized.strip()

def generate_timestamp() -> str:
    """
    Generates a timestamp string for file naming.
    
    Returns:
        str: Timestamp string
    """
    return datetime.now().strftime('%Y%m%d_%H%M%S')

def validate_ai_response(response_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Validates AI response data structure.
    
    Args:
        response_data (dict): AI response data
        
    Returns:
        dict: Validation result
    """
    result = {
        'valid': True,
        'errors': [],
        'warnings': []
    }
    
    # Check required structure
    if not isinstance(response_data, dict):
        result['valid'] = False
        result['errors'].append("Response must be a dictionary")
        return result
    
    # Check for required keys
    required_keys = ['columns', 'rows']
    for key in required_keys:
        if key not in response_data:
            result['valid'] = False
            result['errors'].append(f"Missing required key: {key}")
    
    # Validate columns
    if 'columns' in response_data:
        if not isinstance(response_data['columns'], list):
            result['valid'] = False
            result['errors'].append("Columns must be a list")
    
    # Validate rows
    if 'rows' in response_data:
        if not isinstance(response_data['rows'], list):
            result['valid'] = False
            result['errors'].append("Rows must be a list")
        else:
            # Check each row
            for i, row in enumerate(response_data['rows']):
                if not isinstance(row, dict):
                    result['warnings'].append(f"Row {i} is not a dictionary")
    
    return result

def log_processing_step(step: str, data: Any = None, level: str = 'info'):
    """
    Logs processing steps for debugging and monitoring.
    
    Args:
        step (str): Step description
        data (any, optional): Associated data
        level (str): Log level
    """
    message = f"Processing step: {step}"
    if data:
        message += f" | Data: {str(data)[:100]}..."  # Truncate long data
    
    if level == 'error':
        logger.error(message)
    elif level == 'warning':
        logger.warning(message)
    else:
        logger.info(message)

def create_error_response(error_message: str, error_code: str = None) -> Dict[str, Any]:
    """
    Creates a standardized error response.
    
    Args:
        error_message (str): Error message
        error_code (str, optional): Error code
        
    Returns:
        dict: Error response
    """
    response = {
        'success': False,
        'error': error_message,
        'timestamp': datetime.now().isoformat()
    }
    
    if error_code:
        response['error_code'] = error_code
    
    return response

def create_success_response(data: Any = None, message: str = "Success") -> Dict[str, Any]:
    """
    Creates a standardized success response.
    
    Args:
        data (any, optional): Response data
        message (str): Success message
        
    Returns:
        dict: Success response
    """
    response = {
        'success': True,
        'message': message,
        'timestamp': datetime.now().isoformat()
    }
    
    if data is not None:
        response['data'] = data
    
    return response 