# Workflow Examples and Screenshots

This document provides detailed explanations of the system's workflow through visual examples and screenshots.

## 1. Workflow Diagram Overview

The system follows a structured workflow that processes images and generates cost sheets automatically.

### Workflow Components:

#### **START Node**
- **Icon**: Blue house icon
- **Function**: Initiates the chat process
- **Timing**: 36.882 ms
- **Description**: Entry point where users upload images and begin the analysis process

#### **OCR & OBJECT ANALYSIS Node**
- **Icon**: Purple eye icon with chat, vision, document, and chart icons
- **Function**: Performs optical character recognition and object analysis
- **Timing**: 18.821 seconds
- **Processing**: 82,535 tokens
- **Description**: Analyzes uploaded images to extract text and identify objects

#### **KNOWLEDGE RETRIEVAL Node**
- **Icon**: Green book icon
- **Function**: Accesses pricing database
- **Timing**: 2.672 seconds
- **Data Source**: Price-ncc-doc.csv
- **Description**: Retrieves relevant pricing information from the knowledge base

#### **PRICE-KNOWLEDGE Node**
- **Icon**: Purple eye icon with analysis tools
- **Function**: Processes pricing information
- **Timing**: 29.269 seconds
- **Processing**: 313,620 tokens
- **Description**: Applies pricing logic and calculates costs

#### **SEND TO CREATE EXCEL Node**
- **Icon**: Purple HTTP icon
- **Function**: Sends data to Windmill API
- **Timing**: 609.100 ms
- **Method**: POST request
- **URL**: `https://ncc-windmill.qsncc.com/api/w/analyst-hub/jobs/run_wait_result/f/u/wipasana/grand_flow`
- **Retry Policy**: 3 times on failure
- **Description**: Generates professional Excel cost sheet

#### **CONVERT TO URL FILE Node**
- **Icon**: Blue code icon (`</>`)
- **Function**: Converts Excel to downloadable URL
- **Timing**: 72.316 ms
- **Description**: Prepares file for user download

#### **ANSWER 7 Node**
- **Icon**: Orange question mark icon
- **Function**: Returns final result
- **Timing**: 14.170 ms
- **Description**: Provides download link to the user

## 2. Chat Interface Screenshot

### User Experience Flow:

#### **Initial Greeting**
- **Bot Message**: "Hello! You can upload images or Exhibition Booth files here!"
- **Interface**: Clean, welcoming chat interface
- **Features**: File upload capability with visual indicators

#### **User Interaction**
- **Upload Process**: Users can upload multiple images simultaneously
- **File Types**: Supports various image formats
- **Visual Feedback**: Shows uploaded files with preview

#### **Processing Status**
- **Real-time Updates**: "Workflow Process >" with spinning indicator
- **Progress Tracking**: Shows processing steps in real-time
- **Status Indicators**: Visual cues for ongoing processes

#### **Interface Controls**
- **Stop Response**: Button to halt processing if needed
- **Input Field**: "Talk to Bot" for additional communication
- **Attachment Support**: Paperclip icon for file uploads
- **Send Button**: Blue circular button with airplane icon
- **Features Enabled**: Shows active system capabilities

## 3. Process Execution Screenshot

### Detailed Process Breakdown:

#### **Workflow Process Header**
- **Status**: Green checkmark indicating successful completion
- **Title**: "Workflow Process" with success indicator
- **Overall Status**: All 7 steps completed successfully

#### **Step-by-Step Execution**

1. **START** (36.882 ms)
   - **Status**: ✅ Completed
   - **Function**: Process initialization
   - **Performance**: Very fast execution

2. **OCR & OBJECT ANALYSIS** (18.821 s)
   - **Status**: ✅ Completed
   - **Processing**: 82,535K tokens
   - **Function**: Image analysis and text extraction
   - **Performance**: Most time-consuming step (37% of total)

3. **KNOWLEDGE RETRIEVAL** (2.672 s)
   - **Status**: ✅ Completed
   - **Function**: Database access
   - **Performance**: Efficient data retrieval

4. **PRICE-KNOWLEDGE** (29.269 s)
   - **Status**: ✅ Completed
   - **Processing**: 313,620K tokens
   - **Function**: Price calculation and validation
   - **Performance**: Second most time-consuming step (57% of total)

5. **HTTP SEND TO CREATE EXCEL** (609.100 ms)
   - **Status**: ✅ Completed
   - **Function**: API communication
   - **Performance**: Fast external service call

6. **CONVERT TO URL FILE** (72.316 ms)
   - **Status**: ✅ Completed
   - **Function**: File conversion
   - **Performance**: Very fast processing

7. **ANSWER 7** (14.170 ms)
   - **Status**: ✅ Completed
   - **Function**: Result delivery
   - **Performance**: Instant response

#### **Output Files**
- **Generated File**: Cost-Sheet.xlsx (with thumbs-up icon)
- **Citations**: Price-ncc-doc.csv (data source reference)

## 4. Sample Output Screenshot

### Generated Cost Sheet Analysis:

#### **Header Section**
- **Title**: "Cost Sheet" prominently displayed
- **Project**: "Association for Women's Rights in Development (AWID) 2024"
- **Date**: 07/08/2025
- **Total Cost**: 1,650.00 (displayed in F8)
- **Note**: "ราคานี้เป็นราคาประเมิณ" (This price is an estimate)

#### **Component Categories**

1. **Structure Section**
   - **Items**: Maxima structure, Lockable cabinet
   - **Pricing**: 200.00 and 950.00 respectively
   - **Total**: 1,150.00
   - **Features**: Professional formatting with subtotals

2. **Furniture & Plant Section**
   - **Items**: Carpet, Stool bar, Glass Round table, Chair, Waste paper basket, Plant pot
   - **Pricing**: Various unit prices (100-200 range)
   - **Total**: 500.00
   - **Features**: Comprehensive furniture listing

3. **Graphic Section**
   - **Items**: Printed Graphic (multiple instances)
   - **Pricing**: Not specified in sample
   - **Total**: 0.00
   - **Features**: Custom graphics and signage

4. **Electrical Section**
   - **Items**: Spotlight LED, Socket
   - **Pricing**: Not specified in sample
   - **Total**: 0.00
   - **Features**: Lighting and power components

#### **Data Structure**
- **Columns**: Type, Component, Descriptions, W/L/H, Quantity, Unit, Unit Price, Amounts, Remarks
- **Formatting**: Professional grid layout with proper alignment
- **Calculations**: Automatic total calculations
- **References**: "Knowledge Retrieval" remarks indicating data source

## 5. Key Performance Insights

### **Processing Efficiency**
- **Total Time**: ~51 seconds for complete workflow
- **AI Processing**: 94% of total time (48 seconds)
- **File Generation**: 6% of total time (3 seconds)
- **Success Rate**: 100% completion rate

### **Token Processing**
- **OCR Analysis**: 82,535 tokens (21% of total)
- **Price Processing**: 313,620 tokens (79% of total)
- **Total Tokens**: 396,155 tokens processed

### **System Reliability**
- **Error Handling**: Comprehensive retry mechanisms
- **Data Validation**: Multi-stage verification process
- **Output Quality**: Professional formatting and accuracy

## 6. User Experience Highlights

### **Simplicity**
- **One-Click Upload**: Easy image upload process
- **Real-Time Feedback**: Live progress updates
- **Instant Results**: Quick file generation and download

### **Professional Output**
- **Structured Format**: Organized cost sheet layout
- **Detailed Breakdown**: Comprehensive component listing
- **Professional Styling**: Ready-to-use business documents

### **Reliability**
- **Consistent Results**: Standardized processing
- **Data Accuracy**: AI-powered precision
- **Secure Processing**: Protected data handling

---

*These examples demonstrate the system's capability to transform simple image uploads into professional cost estimation documents through intelligent AI processing.* 