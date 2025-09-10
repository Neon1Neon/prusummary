# Tax Data Processor

## Overview
A Flask web application that processes Excel files containing tax data and generates monthly tally reports. The application extracts data from uploaded Excel files, processes it according to specific tax data columns, and creates comprehensive reports showing processed tasks by month and processor.

## Recent Changes (September 10, 2025)
- Set up Flask web application with file upload functionality
- Created HTML templates with Bootstrap styling for professional UI
- Integrated existing processor.py logic into Flask routes
- Configured Flask to run on 0.0.0.0:5000 for Replit environment
- Set up workflow for automatic server management
- Added file processing with error handling and user feedback

## Project Architecture
- **app.py**: Main Flask application with file upload and processing routes
- **processor.py**: Original data processing logic (now integrated into app.py)
- **templates/**: HTML templates for the web interface
  - **base.html**: Base template with Bootstrap styling
  - **index.html**: File upload interface
  - **results.html**: Results display with download functionality
- **pyproject.toml**: Python dependencies managed by uv
- **uv.lock**: Dependency lock file

## Key Features
- Multiple Excel file upload support
- Date extraction from filename (format: "Summary [Month Day, Year].xlsx")
- Data processing with 45 predefined tax data columns
- Monthly tally generation grouped by processor and task
- Results display in formatted HTML table
- CSV download functionality
- Bootstrap-styled responsive interface
- Error handling and user feedback via Flash messages

## Dependencies
- **Flask**: Web framework
- **pandas**: Data processing
- **openpyxl**: Excel file reading
- **Werkzeug**: File handling utilities

## Configuration
- **Host**: 0.0.0.0 (required for Replit proxy)
- **Port**: 5000 (Replit standard)
- **Debug Mode**: Enabled for development
- **Max File Size**: 16MB per upload

## Deployment Settings
- **Target**: Autoscale (stateless web application)
- **Run Command**: python app.py