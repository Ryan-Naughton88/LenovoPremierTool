# Lenovo Premier Tool

## Overview

The **Lenovo Premier Tool** is a Python script utilizing the `customtkinter` library and `win32com.client` for creating and managing templates and emails related to Lenovo Premier Technical Support. The tool provides a graphical user interface (GUI) for users to input customer and case information, as well as generate various templates and emails for different scenarios.

## The GUI

![Lenovo Premier Tool Screenshot](Images/Premier-Tool-Screenshot)

## Features

1. **Customer Info:**
   - Input fields for customer name, email, phone, and address.

2. **Machine Info:**
   - Input fields for model, serial number, and MTM (Machine Type Model).

3. **Case Info:**
   - Input fields for case number, work order number, and current work order number.
   - Dropdown for selecting the service type (Onsite, Depot, Parts Only).

4. **Buttons:**
   - **Reset Fields:** Clears all input fields.
   - **Notes Template:** Copies a predefined notes template to the clipboard.
   - **BitLocker Email:** Generates an email template for BitLocker-related instructions.
   - **WO Email:** Generates an email template for Work Order-related information.
   - **WO Template:** Copies a predefined Work Order template to the clipboard.
   - **Depot Template:** Copies a predefined Depot template to the clipboard.

## Usage

1. **Customer Information:**
   - Fill in the relevant customer details in the input fields.

2. **Machine Information:**
   - Provide details about the machine, including model, serial number, and MTM.

3. **Case Information:**
   - Enter case number, work order number, and current work order number.
   - Select the service type from the dropdown.

4. **Buttons:**
   - Use the buttons to perform various actions, such as copying templates or generating emails.

5. **Notes Textbox:**
   - Displays the Premier Notes Template.
   - Can be used to view or edit the template.

## Email Generation

- The tool includes functionality to generate emails for BitLocker-related instructions and Work Order-related information. These emails are opened using Microsoft Outlook.

## Templates

- **Notes Template:** A predefined template for taking notes during customer interactions.
- **WO Template:** A predefined template for Work Order-related information.
- **Depot Template:** A predefined template for Depot-related information.

## Requirements

- `customtkinter` library
- `win32com.client` library
- Python 3.x

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/Ryan-Naughton88/lenovo-premier-tool.git
   cd lenovo-premier-tool
   ```

2. Install required libraries:
   ```bash
   pip install customtkinter
   ```

3. Run the script:
   ```bash
   python lenovo_premier_tool.py
   ```

## Notes

- The tool uses `customtkinter` for the GUI and `win32com.client` for interacting with Microsoft Outlook.
- Ensure that Microsoft Outlook is configured on your system for the email generation functionality.

## License

This project is licensed under the [MIT License](LICENSE). Feel free to use, modify, and distribute the code as per the terms of the license.

---
