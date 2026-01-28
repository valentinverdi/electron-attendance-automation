# electron-attendance-automation
Desktop application built with Electron to automate attendance sheet generation and processing using a local database and Microsoft Excel integration.
The project was originally developed for a real client and later adapted for demonstration and technical evaluation purposes.

Features
  - Desktop application for managing and generating attendance sheets.
  - Local data persistence using SQLite.
  - Automated modification of Excel files through programmatic integration.
  - Execution of Excel workflows directly from the desktop application.
  - Automatic generation of PDF files from Excel sheets.
  - Direct printing of attendance sheets from the desktop application and Excel.
  - Workflow automation designed to reduce manual steps and improve efficiency.

Technologies Used
  - Electron
  - JavaScript
  - HTML / CSS
  - SQLite3 (local database)
  - Microsoft Excel
  - VBA (Visual Basic for Applications) for Excel automation

Electron - Excel Integration      
  - This application integrates with Microsoft Excel to automate attendance-related workflows.
  - The Electron application:
      * Modifies Excel files programmatically.
      * Triggers Excel execution.
      * Controls workflow execution through predefined logic.
  - Excel automation is handled using VBA macros, which:
      * Process attendance data.
      * Configure print areas and page layout.
      * Export spreadsheets as PDF files.
      * Send documents directly to the printer.
      * Automatically close Excel once the process is completed.
  - This integration enables a fully automated workflow from the desktop application to document generation and output.

Project Structure
  - This repository contains only the relevant and representative source code for demonstration purposes.
  * electron-attendance-automation/
  * ├── src/
  * │   ├── main/
  * │   ├── renderer/
  * │   └── /
  * ├── excel/
  * │   └── vba/
  * │       └── attendance_macros.bas
  * └── README.md

Excel Automation (VBA)
  - The attendance_macros.bas file included in this repository contains a selected subset of VBA macros used by the application.
  - These macros were chosen to demonstrate:
    * Excel workflow automation.
    * PDF generation and printing.
    * Interaction between the Electron application and Excel.
  - Not all original macros are included, and business-specific logic has been intentionally excluded to preserve client confidentiality.

Preview
Demo video showing the application workflow, including Excel execution and PDF generation:
https://youtu.be/OhM0ZOh3MU8

Confidentiality Notice
  - As this project was developed for a real client, the following elements have been removed:
      * Private credentials.
      * Sensitive business data.
  - This repository focuses on showcasing technical integration, automation logic and project structure.

Project Status
  - Fully functional production project.
  - Public version adapted for demonstration and technical evaluation.

Author

Valentín Germán Verdi
Advanced Student of Computer Engineering
