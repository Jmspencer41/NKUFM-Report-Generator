# NKU Facilities Report Generator

A specialized tool designed for Northern Kentucky University's Facilities Management Department to streamline project reporting by converting Excel data from the project planning database into professionally formatted PDF reports for project meetings.

## Overview

This application automates the time-consuming process of manually creating project reports by:
- Reading Excel files exported from NKU's project planning database
- Processing and formatting the data according to department standards
- Generating clean, professional PDF reports ready for project meetings
- Ensuring consistent formatting and presentation across all reports

## Download

[![Download Latest Release](https://img.shields.io/github/v/release/Jmspencer41/Report-Generator?label=Download&style=for-the-badge)](https://github.com/Jmspencer41/NKUFM-Report-Generator/releases/tag/v1.0.0)

### System Requirements
- Windows 10 or later
- Java Runtime Environment (JRE) 11 or higher
- Microsoft Excel (for viewing source files)
- PDF viewer (Adobe Reader, etc.)

## Features

- **Excel File Processing**: Seamlessly reads and processes Excel files from NKU's project planning database
- **Automated Formatting**: Converts raw data into department-standard report format
- **PDF Generation**: Creates professional PDF reports suitable for meetings and presentations
- **User-Friendly Interface**: Simple, intuitive interface designed for non-technical users
- **Error Handling**: Built-in validation to ensure data integrity and proper formatting
- **Batch Processing**: Handle multiple projects or reports efficiently

## How to Use

1. **Export Data**: Export your project data from the NKU project planning database to an Excel file
2. **Launch Application**: Run the Report Generator executable
3. **Select File**: Browse and select your Excel file containing project data
4. **Generate Report**: Click the generate button to create your PDF report
5. **Save & Share**: The PDF report will be saved to your specified location, ready for your project meeting

## Input File Requirements

The Excel file should contain the following standard columns from the NKU project planning database:
- Project ID/Number
- Project Name
- Status
- Budget Information
- Timeline/Dates
- Responsible Parties
- Project Description

*Note: The application is configured to work with the standard export format from NKU's project planning database.*

## Output

The generated PDF report includes:
- Executive summary of projects
- Detailed project status information
- Budget summaries and financial data
- Timeline and milestone information
- Professional formatting suitable for presentation

## Installation

1. Download the latest release from the link above
2. Run the `.exe` file
3. Follow the installation prompts
4. Launch the application from your desktop or start menu

## Support

For technical support or questions regarding this application:
- **Internal Users**: Contact NKU Facilities Management IT Support
- **Development Issues**: Create an issue in this repository

## Technical Information

- **Built with**: Java and JavaFX
- **PDF Library**: Apache PDFBox
- **Excel Processing**: Apache POI
- **Target Platform**: Windows (64-bit)

## Version History

See the [Releases](https://github.com/Jmspencer41/Report-Generator/releases) page for detailed version history and updates.

## License

This software is developed specifically for Northern Kentucky University's Facilities Management Department. Internal use only.

---

**Developed for NKU Facilities Management Department**  