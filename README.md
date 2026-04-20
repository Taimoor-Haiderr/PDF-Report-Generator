#  PDF Report Generator

A Python-based GUI application to generate **secure, professional PDF reports** for Student and Company data using structured formatting, charts, and encryption.


# Project Overview

This project simulates a real-world reporting system used in organizations to manage and generate structured reports. 
It allows users to input, load, preview, and export data into professionally formatted PDF documents.


# Features

# Data Management
- Add records manually via GUI
- Load data from:
  - CSV files (Students)
  - JSON files (Companies)
- Preview all records before generating reports

#PDF Report Generation
- Clean and professional layout
- Dynamic table formatting
- Auto-adjusted column widths
- Date & time stamping
- Optional company logo support

# Data Visualization
- Automatic chart generation using Matplotlib
- Performance summary visualization included in PDF

# Security Features
- Password-protected PDF
- Restricted permissions:
  -  Copy disabled
  -  Edit disabled
  -  Print disabled
  -  Default password (1234).

# File Management
- Unique filenames using timestamps
- Reports saved in `Reports/` directory
- Report history tracking inside GUI


# GUI Interface

Built using Tkinter with:
- Multi-tab layout:
  - Add Data
  - Load Data
  - Preview Records
  - Generate PDF
  - History
- Clean and user-friendly design


# Technologies Used

- Python 
- Tkinter (GUI)
- ReportLab (PDF generation)
- Matplotlib (Charts)
- CSV / JSON handling

#Author

Developed by: Taimoor Haider
Field: Cybersecurity & Python Development
