# Certificate Automation Script for Excel Training

This project is a Google Apps Script designed to automate the generation and emailing of certification documents for **Excel Training** based on employee performance data in a Google Sheet. The script integrates with Google Sheets, Google Drive, and Gmail to streamline the certification process.

---

## Features

1. **Automatic Certificate Generation**:
   - Certificates are generated when a score cell is highlighted green in the Google Sheet.
   - Supports three certification levels: Basic, Intermediate, and Advanced.

2. **Email Integration**:
   - Automatically emails the generated certificates to employees.
   - Includes a customizable email template with a professional signature and branding.

3. **Duplicate Prevention**:
   - Checks for existing certificates in Google Drive to avoid duplicates.

4. **Customizable Settings**:
   - Enable or disable automatic certificate generation.
   - Enable or disable automatic email sending.

5. **Bulk Certificate Generation**:
   - Generate certificates for all employees or specific certification levels.

6. **User-Friendly Menu**:
   - Adds a custom menu to the Google Sheet for easy access to script features.

---

## Screenshots

### 1. **Certificate Auto-Generation Settings**
![Certificate Auto-Generation Settings](./images/certificate-auto-generation-settings.png)

### 2. **Email Auto-Sending Settings**
![Email Auto-Sending Settings](./images/email-auto-sending-settings.png)

### 3. **Custom Menu for Certificates**
![Custom Menu for Certificates](./images/custom-menu-for-certificates.png)

### 4. **Basic Questions Sheet**
![Basic Questions Sheet](./images/basic-questions-sheet.png)

### 5. **Score Tracker for Active Employees**
![Score Tracker for Active Employees](./images/score-tracker-for-active-employees.png)

### 6. **Generated Email Example**
![Generated Email with Footer](./images/generated-email-with-footer.png)

### 7. **Generated Certificates in Google Drive**
![Generated Certificates in Google Drive](./images/generated-certificates-in-google-drive.png)

---

## File Structure

```
Automation/
├── Automation.js
├── images/
    ├── certificate-auto-generation-settings.png
    ├── email-auto-sending-settings.png
    ├── custom-menu-for-certificates.png
    ├── basic-questions-sheet.png
    ├── score-tracker-for-active-employees.png
    ├── generated-email-with-footer.png
    ├── generated-certificates-in-google-drive.png
```

### Key File:
- **Automation.js**: Main script file containing the logic for certificate generation, email sending, and menu integration.

---

## Setup Instructions

### 1. **Prerequisites**
   - A Google Workspace account with access to Google Sheets, Google Drive, and Gmail.
   - A Google Sheet with employee data, including names, scores, and email addresses.

### 2. **Open Google Apps Script**
   - Open the Google Sheet where you want to use this script.
   - Navigate to `Extensions > Apps Script`.

### 3. **Add the Script**
   - Copy the contents of `Automation.js` and paste it into the Apps Script editor.

### 4. **Set Up Google Drive Folders**
   - Create folders in Google Drive for storing certificates:
     - Basic Certificates
     - Intermediate Certificates
     - Advanced Certificates
   - Update the folder IDs in the script:
     ```javascript
     const BASIC_TEMPLATE_ID = 'YOUR_BASIC_TEMPLATE_ID';
     const INTERMEDIATE_TEMPLATE_ID = 'YOUR_INTERMEDIATE_TEMPLATE_ID';
     const ADVANCED_TEMPLATE_ID = 'YOUR_ADVANCED_TEMPLATE_ID';
     ```

### 5. **Set Up Email Templates**
   - Customize the email subject and body in the script:
     ```javascript
     const EMAIL_SUBJECT_TEMPLATE = '%s Certification Training Completed';
     const EMAIL_BODY_TEMPLATE = `Dear %s,

     Congratulations on successfully completing the Excel %s Certification Training.

     We are pleased to present your official certification document, which is attached to this email. This certification recognizes your proficiency with Microsoft Excel and validates your expertise at the %s level.

     Your achievement demonstrates both your commitment to developing valuable data analysis skills and your investment in expanding your professional capabilities. We encourage you to add this Excel certification to your professional profile and resume.

     If you have any questions regarding your certification or wish to explore additional Excel training opportunities, please do not hesitate to contact us.

     Best regards,
     Training Certification Team`;
     ```

### 6. **Authorize the Script**
   - Save the script and run any function (e.g., `onOpen`) to trigger the authorization process.
   - Grant the necessary permissions.

### 7. **Enable Triggers**
   - Set up a trigger for the `onEdit` function to monitor changes in the Google Sheet:
     - Go to `Triggers` in the Apps Script editor.
     - Add a new trigger for `onEdit`.

---

## Usage

### 1. **Custom Menu**
   - After setup, a new menu called `Certificates` will appear in the Google Sheet.
   - Use this menu to:
     - Generate certificates for all employees or specific levels.
     - Enable/disable automatic certificate generation.
     - Enable/disable automatic email sending.

### 2. **Highlight Cells**
   - Highlight a score cell green to trigger certificate generation.

### 3. **View Logs**
   - Use the `Logger.log` statements in the script to debug or monitor the process.

---

## Customization

### 1. **Email Footer**
   - Modify the footer in the email template to match your organization's branding.

### 2. **Certificate Templates**
   - Update the Google Doc templates for each certification level with placeholders:
     - `<<NAME>>` for the employee's name.
     - `<<DATE>>` for the certification date.

### 3. **Menu Options**
   - Add or remove menu options in the `onOpen` function as needed.

---

## Troubleshooting

### 1. **Certificates Not Generating**
   - Ensure the `AUTO_CERT_GENERATION_ENABLED` property is set to `true`.
   - Verify the folder IDs and template IDs in the script.

### 2. **Emails Not Sending**
   - Ensure the `AUTO_EMAIL_ENABLED` property is set to `true`.
   - Check the email addresses in the Google Sheet.

### 3. **Duplicate Certificates**
   - Ensure the script has access to the correct Google Drive folders.

---

## License

This project is licensed under the MIT License. You are free to use, modify, and distribute this script.

---

## Contact

For questions or support, please contact the Training Certification Team at `reysan.aretex@gmail.com`.