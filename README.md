# T-MOBILE - START/STOP Mail 😋

## Overview
`T-MOBILE - START/STOP Mail 😋` is a Python script designed to automate the creation and sending of "START" and "STOP" emails using Microsoft Outlook. This script helps in maintaining daily logs by sending emails with predefined subjects and bodies based on user input.

## Features
- Automatically generates email subject lines based on the current date and user input.
- Prompts user for additional details to include in the email body when sending a "STOP" email.
- Configurable recipients, CC list, and focus settings.
- Saves the created email as a draft in Outlook.
- Provides visual feedback in the console using `colorama`.

## Requirements
- Windows OS (required for `win32com.client` to interact with Outlook)
- Python 3.x
- `win32com.client`
- `inquirer`
- `colorama`
- `datetime`
- `json`

## Installation
1. Ensure you have Python 3.x installed on your system.
2. Install the required Python packages:
    ```bash
    pip install pypiwin32 inquirer colorama
    ```

## Configuration
Create a `config.json` file in the same directory as the script with the following structure:
```json
{
    "to": "recipient@example.com",
    "title": "Daily Report",
    "cc": ["cc1@example.com", "cc2@example.com"],
    "focus": true
}
```


