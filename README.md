# Automated Invoice
This Python script takes monthly totals for, in this case, translation fees from an excel table. It will write and send an email to the client, the email message states the total fee for that month, and attaches the payments table file and an invoice file.

## Table of Contents
- [Overview](#overview)
- [Usage](#usage)
- [Files](#files)
- [Email Content](#email-content)
- [Dependencies](#dependencies)
- [Important Notes](#important-notes)
- [Author](#author)
- [License](#license)

## Overview
The script performs the following tasks:

- Determines the correct year and month for the invoice based on the current date.
- Retrieves payment data from an Excel file and calculates the total payment.
- Updates the Payments Excel file with the total payment.
- Creates an invoice using a template, updating relevant cells with invoice details.
- Saves the invoice as a new Excel document.
- Sends an email with the invoice and payment details as attachments.

## Usage
1. Make sure you have the required dependencies installed:

   ```bash
   pip install openpyxl
   ```

2. Update the following information in the Python script:

Set your email credentials (email_sender, email_password, email_receiver).
Customize the file names and paths for the Payments Excel file and the invoice template.
  
3. Run the script:

   ```bash
   python monthly_invoice_generator.py
   ```

Check your email for the generated invoice and payment details.

## Files
`monthly_invoice_generator.py`: Python script for generating and sending invoices.
`202X年X月翻訳.xlsx`: Invoice template.
`{invoice_year} Payments.xlsx`: Excel file containing payment data.

## Email Content
The email includes the following information:

Subject: 今月の請求書 (This month's invoice)
Body: A message with details about the invoice and a statement of completion.

## Dependencies
[openpyxl](https://pypi.org/project/openpyxl/)

## Important Notes
The script is configured to use a Gmail SMTP server. Ensure that your email provider allows SMTP access and adjust server details if necessary.

## Author
Alex McKinley

## License
This project is licensed under the [MIT License](LICENSE).
