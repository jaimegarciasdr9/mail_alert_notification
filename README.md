# Email Alert Notification Script

## Overview
This Python script automates the process of sending email notifications regarding employees without a DNI (document identification number) in a company's internal system. It connects to a database, retrieves relevant employee data, generates an Excel file listing the employees without a DNI, and sends the report via email to a designated recipient.

The script performs the following key actions:

- Database Connection: Connects to an internal database and retrieves employee data.

- Data Processing: Extracts a list of employees without a DNI and generates an Excel report.

- Email Notification: Sends an automated email with the attached Excel report to a specified recipient.

## Key Functions
1. mail_alert_notification()

This is the main function that orchestrates the entire workflow:

 * Connects to the database.
   
 * Queries for employees missing a DNI.
   
 * Exports the data to an Excel file.
   
 * Sends an email with the generated Excel report as an attachment.

2. automatic_email_send(receiver, subject, message, file_attach)

This function is responsible for sending the automated email:

  * Takes parameters such as the receiver's email, subject, message body, and the file to be attached.
    
  * Constructs an HTML-formatted email and sends it via an SMTP server.

## Requirements
To run this script, ensure you have the following Python packages installed:

pip install pandas xlsxwriter smtplib

You may need to configure additional email settings depending on your SMTP provider.
