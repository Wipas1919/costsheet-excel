"""
EC - AI Cost Estimation System Demo

This file demonstrates how to use the cost estimation system.
All sensitive data has been replaced with sample data.
"""

import json
from excel import main as generate_excel
from utils import validate_ai_response, log_processing_step
from config import get_config

def demo_ai_analysis():
    """
    Demonstrates AI analysis workflow with sample data.
    """
    print("=== EC AI Cost Estimation System Demo ===\n")
    
    # Sample AI response (simulating what would come from Gemini)
    sample_ai_response = {
        "columns": [
            "list_id", "Component", "Description", "W", "L", "H",
            "Quantity", "Unit", "price_per_unit", "total_cost", "remark"
        ],
        "rows": [
            {
                "list_id": "100-01-01",
                "Component": "Flooring",
                "Description": "Carpet flooring",
                "W": "10.0",
                "L": "5.0",
                "H": "0.1",
                "Quantity": "50.0",
                "Unit": "sqm",
                "price_per_unit": "150.00",
                "total_cost": "7500.00",
                "remark": "Sample flooring"
            },
            {
                "list_id": "100-02-01",
                "Component": "Structure",
                "Description": "Wall panel",
                "W": "3.0",
                "L": "2.5",
                "H": "2.5",
                "Quantity": "15.0",
                "Unit": "sqm",
                "price_per_unit": "200.00",
                "total_cost": "3000.00",
                "remark": "Sample structure"
            },
            {
                "list_id": "102-01-01",
                "Component": "Graphic",
                "Description": "Printed vinyl",
                "W": "2.0",
                "L": "1.5",
                "H": "-",
                "Quantity": "3.0",
                "Unit": "sqm",
                "price_per_unit": "80.00",
                "total_cost": "240.00",
                "remark": "Sample graphic"
            }
        ]
    }
    
    # Validate AI response
    print("1. Validating AI Response...")
    validation_result = validate_ai_response(sample_ai_response)
    
    if validation_result['valid']:
        print("SUCCESS: AI response validation passed")
        if validation_result['warnings']:
            print(f"WARNING: {validation_result['warnings']}")
    else:
        print(f"ERROR: Validation failed: {validation_result['errors']}")
        return
    
    # Extract columns and rows
    columns = sample_ai_response['columns']
    rows = []
    
    for row_data in sample_ai_response['rows']:
        row = []
        for col in columns:
            row.append(row_data.get(col, '-'))
        rows.append(row)
    
    print(f"SUCCESS: Extracted {len(rows)} rows with {len(columns)} columns")
    
    # Generate Excel file
    print("\n2. Generating Excel Cost Sheet...")
    log_processing_step("Starting Excel generation", f"Rows: {len(rows)}, Columns: {len(columns)}")
    
    result = generate_excel(columns, rows)
    
    if result['success']:
        print(f"SUCCESS: Excel file generated: {result['filename']}")
        print(f"INFO: Message: {result['message']}")
    else:
        print(f"ERROR: Failed to generate Excel: {result['error']}")
    
    return result

def demo_configuration():
    """
    Demonstrates system configuration.
    """
    print("\n=== System Configuration Demo ===\n")
    
    # Get configuration
    config = get_config('development')
    
    print("System Information:")
    print(f"  Name: {config.SYSTEM_NAME}")
    print(f"  Version: {config.VERSION}")
    print(f"  AI Model: {config.AI_MODEL}")
    print(f"  Vision Detail: {config.VISION_DETAIL}")
    
    print("\nComponent Types:")
    for code, name in config.get_component_types().items():
        print(f"  {code}: {name}")
    
    print("\nValid Units:")
    for unit in config.VALID_UNITS:
        print(f"  - {unit}")
    
    print("\nSecurity Features:")
    for feature, enabled in config.SECURITY.items():
        status = "ENABLED" if enabled else "DISABLED"
        print(f"  {feature}: {status}")

def demo_workflow():
    """
    Demonstrates the complete workflow.
    """
    print("\n=== Complete Workflow Demo ===\n")
    
    workflow_steps = [
        "1. User uploads exhibition booth images",
        "2. AI analyzes images and extracts components",
        "3. System classifies components by type",
        "4. Dimensions are extracted and validated",
        "5. Quantities are calculated based on component type",
        "6. Prices are retrieved from knowledge base",
        "7. Total costs are calculated",
        "8. Excel cost sheet is generated",
        "9. File is uploaded to cloud storage",
        "10. Download link is provided to user"
    ]
    
    print("Workflow Steps:")
    for step in workflow_steps:
        print(f"  {step}")
    
    print("\nData Flow:")
    print("  Images → AI Analysis → Structured Data → Price Calculation → Excel Generation → Cloud Storage")

def demo_error_handling():
    """
    Demonstrates error handling scenarios.
    """
    print("\n=== Error Handling Demo ===\n")
    
    from utils import create_error_response, create_success_response
    
    # Sample error scenarios
    error_scenarios = [
        ("Invalid image format", "INVALID_FORMAT"),
        ("AI analysis failed", "AI_ERROR"),
        ("Missing required data", "MISSING_DATA"),
        ("Network connection error", "NETWORK_ERROR")
    ]
    
    print("Error Response Examples:")
    for error_msg, error_code in error_scenarios:
        error_response = create_error_response(error_msg, error_code)
        print(f"  {error_code}: {error_msg}")
        print(f"    Response: {json.dumps(error_response, indent=4)}")
    
    print("\nSuccess Response Example:")
    success_response = create_success_response(
        data={"filename": "Cost_Sheet_20241201_143022.xlsx"},
        message="Cost sheet generated successfully"
    )
    print(f"  Response: {json.dumps(success_response, indent=4)}")

if __name__ == "__main__":
    """
    Run the complete demonstration.
    """
    try:
        # Run all demos
        demo_ai_analysis()
        demo_configuration()
        demo_workflow()
        demo_error_handling()
        
        print("\n" + "="*50)
        print("🎉 Demo completed successfully!")
        print("="*50)
        
    except Exception as e:
        print(f"\n❌ Demo failed with error: {str(e)}")
        print("This is expected behavior for demonstration purposes.") 