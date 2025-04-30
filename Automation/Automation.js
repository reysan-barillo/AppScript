const BASIC_TEMPLATE_ID = '1DceRNuLO-wUU1JeZoRzfVbZfX6skwDhiquvhx0ikflM';
const INTERMEDIATE_TEMPLATE_ID = '1t-dHUBljFCGTVSOXZur7beHaQovhPiAMxb_gak0bbYI';
const ADVANCED_TEMPLATE_ID = '1y6UhwHwKcBqtbzNyoxiU0JP4cHNxy5dbXizuLwUx5RE';
const DEST_FOLDER_ID = '1TZbpNSdG-IqAlSMnrvVn4PpdWB12S0yH';
const SHEET_NAME = 'Score Tracker For Active Employee';
const NAME_COLUMN = 2;
const BASIC_SCORE_COLUMN = 5;
const INTERMEDIATE_SCORE_COLUMN = 6;
const ADVANCED_SCORE_COLUMN = 7;

// Define color codes for Google Sheets
const GREEN_COLOR = '#93c47d';
const RED_COLOR = '#e06666';  

// Email settings
const EMAIL_SUBJECT_TEMPLATE = '%s Certification Training Completed';
const EMAIL_BODY_TEMPLATE = `Dear %s,

Congratulations on successfully completing the Excel %s Certification Training.

We are pleased to present your official certification document, which is attached to this email. This certification recognizes your proficiency with Microsoft Excel and validates your expertise at the %s level.

Your achievement demonstrates both your commitment to developing valuable data analysis skills and your investment in expanding your professional capabilities. We encourage you to add this Excel certification to your professional profile and resume.

If you have any questions regarding your certification or wish to explore additional Excel training opportunities, please do not hesitate to contact us.

Best regards,
Training Certification Team`;

function onEdit(e) {
  const isAutoGenerationEnabled = PropertiesService.getScriptProperties().getProperty('AUTO_CERT_GENERATION_ENABLED') === 'true';
  const isAutoEmailEnabled = PropertiesService.getScriptProperties().getProperty('AUTO_EMAIL_ENABLED') === 'true';

  if (!e || !e.range || e.range.getSheet().getName() !== SHEET_NAME) {
    return;
  }

  const column = e.range.getColumn();
  const row = e.range.getRow();

  if (row <= 1) {
    return;
  }

  // Only process if this is a score column
  if (column !== BASIC_SCORE_COLUMN && 
      column !== INTERMEDIATE_SCORE_COLUMN && 
      column !== ADVANCED_SCORE_COLUMN) {
    return;
  }

  const editedValue = e.range.getValue();
  
  // If cell is completely cleared (not just set to 0), then don't change formatting
  if (editedValue === null || editedValue === "") {
    return; 
  }
  
  // Check if value is not a valid number
  if (isNaN(editedValue)) {
    e.range.setBackground(RED_COLOR); 
    return;
  }

  const score = parseInt(editedValue, 10);
  let examType, scoreColumn, templateId, destFolderId, questionSheetName, statusColumn;

  if (column === BASIC_SCORE_COLUMN) {
    examType = 'Basic';
    scoreColumn = BASIC_SCORE_COLUMN;
    templateId = BASIC_TEMPLATE_ID;
    destFolderId = "1giX-nYnriLX9IemmGpNXHiCtafProbTo";
    questionSheetName = 'Basic Questions';
    statusColumn = 8; 
  } else if (column === INTERMEDIATE_SCORE_COLUMN) {
    examType = 'Intermediate';
    scoreColumn = INTERMEDIATE_SCORE_COLUMN;
    templateId = INTERMEDIATE_TEMPLATE_ID;
    destFolderId = "171I3Ll59dNHCFxhE7wkg3GPxtfwg_fnv";
    questionSheetName = 'Intermediate Questions';
    statusColumn = 9; 
  } else if (column === ADVANCED_SCORE_COLUMN) {
    examType = 'Advanced';
    scoreColumn = ADVANCED_SCORE_COLUMN;
    templateId = ADVANCED_TEMPLATE_ID;
    destFolderId = "1f0XCRnGgmFPkOVsHHilm7B8Z5er3keic";
    questionSheetName = 'Advanced Questions';
    statusColumn = 10; 
  } else {
    return;
  }

  if (score >= 18) {
    e.range.setBackground(GREEN_COLOR);

    const sheet = e.range.getSheet();

    if (!isAutoGenerationEnabled) {
      // Update the status column with "Certificate is available" when auto-generation is disabled
      sheet.getRange(row, statusColumn).setValue('Certificate is available');
      return;
    }

    const name = sheet.getRange(row, NAME_COLUMN).getValue();

    const questionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(questionSheetName);
    if (!questionSheet) {
      SpreadsheetApp.getUi().alert(`${questionSheetName} sheet not found.`);
      return;
    }

    const questionData = questionSheet.getDataRange().getValues();
    let email = null;

    for (let i = 1; i < questionData.length; i++) {
      if (questionData[i][2] === name) {
        email = questionData[i][3];
        break;
      }
    }

    if (!email) {
      SpreadsheetApp.getUi().alert(`Email not found for ${name} in ${questionSheetName} sheet.`);
      return;
    }

    const currentDate = new Date();
    const formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "MMMM d, yyyy");

    try {
      const certificateGenerated = generateSingleCertificate(name, examType, templateId, formattedDate, email, isAutoEmailEnabled);

      if (certificateGenerated) {
        if (!isAutoEmailEnabled) {
          // Update the status column when auto-email is disabled
          sheet.getRange(row, statusColumn).setValue('Certificate is available (Email not sent - auto-email disabled)');
        } else {
          sheet.getRange(row, statusColumn).setValue(`Certificate already sent to ${email}`);
        }
      } else {
        sheet.getRange(row, statusColumn).setValue('Generating Certificate');
      }
    } catch (error) {
      Logger.log(`Error generating certificate: ${error.toString()}`);
      sheet.getRange(row, statusColumn).setValue('Error during certificate generation');
    }
  } else {
    e.range.setBackground(RED_COLOR);
    const sheet = e.range.getSheet();
    sheet.getRange(row, statusColumn).setValue('Failed Score');
  }
}

function processAutoGenerateCertificate() {
  const props = PropertiesService.getScriptProperties();
  const allProps = props.getProperties();
 
  let certificatesGenerated = 0;
  let emailsSent = 0;
 
  // Map to track existing certificates for each folder
  let existingCertificatesByFolder = {
    "1giX-nYnriLX9IemmGpNXHiCtafProbTo": new Map(),
    "171I3Ll59dNHCFxhE7wkg3GPxtfwg_fnv": new Map(),
    "1f0XCRnGgmFPkOVsHHilm7B8Z5er3keic": new Map()  
  };
 
  // Build a list of all existing certificates in each folder to prevent duplicates
  try {
    // Check Basic folder
    const basicFolder = DriveApp.getFolderById("1giX-nYnriLX9IemmGpNXHiCtafProbTo");
    let allFiles = basicFolder.getFiles();
    while (allFiles.hasNext()) {
      const file = allFiles.next();
      existingCertificatesByFolder["1giX-nYnriLX9IemmGpNXHiCtafProbTo"].set(file.getName().toLowerCase(), file.getId());
    }
   
    // Check Intermediate folder
    const intermediateFolder = DriveApp.getFolderById("171I3Ll59dNHCFxhE7wkg3GPxtfwg_fnv");
    allFiles = intermediateFolder.getFiles();
    while (allFiles.hasNext()) {
      const file = allFiles.next();
      existingCertificatesByFolder["171I3Ll59dNHCFxhE7wkg3GPxtfwg_fnv"].set(file.getName().toLowerCase(), file.getId());
    }
   
    // Check Advanced folder
    const advancedFolder = DriveApp.getFolderById("1f0XCRnGgmFPkOVsHHilm7B8Z5er3keic");
    allFiles = advancedFolder.getFiles();
    while (allFiles.hasNext()) {
      const file = allFiles.next();
      existingCertificatesByFolder["1f0XCRnGgmFPkOVsHHilm7B8Z5er3keic"].set(file.getName().toLowerCase(), file.getId());
    }
   
    Logger.log(`Found existing certificates - Basic: ${existingCertificatesByFolder["1giX-nYnriLX9IemmGpNXHiCtafProbTo"].size}, Intermediate: ${existingCertificatesByFolder["171I3Ll59dNHCFxhE7wkg3GPxtfwg_fnv"].size}, Advanced: ${existingCertificatesByFolder["1f0XCRnGgmFPkOVsHHilm7B8Z5er3keic"].size}`);
  } catch (e) {
    Logger.log(`Error building certificate map: ${e.toString()}`);
  }
 
  // Now process certificate generation requests
  for (const key in allProps) {
    if (key.includes('_') && (key.includes('Basic') || key.includes('Intermediate') || key.includes('Advanced'))) {
      try {
        const certData = JSON.parse(allProps[key]);
        const certificateName = `${certData.examType} Certificate - ${certData.name}.pdf`;
        const destFolderId = certData.destFolderId;
        let pdfFileId = null;
        let pdfFile = null;
       
        // Check if certificate already exists before generating
        if (existingCertificatesByFolder[destFolderId].has(certificateName.toLowerCase())) {
          Logger.log(`${certData.examType} Certificate for ${certData.name} already exists in the ${certData.examType} folder, using existing for email`);
          // Get the existing file ID for email attachment
          pdfFileId = existingCertificatesByFolder[destFolderId].get(certificateName.toLowerCase());
          pdfFile = DriveApp.getFileById(pdfFileId);
        } else {
          // Generate the certificate and convert to PDF, cleaning up the Google Doc
          const templateDoc = DriveApp.getFileById(certData.templateId);
          const destFolder = DriveApp.getFolderById(destFolderId);
         
          // Make a copy of the template for the employee's certificate with timestamp to avoid conflicts
          const tempTimestamp = new Date().getTime();
          const tempName = `${certData.examType} Certificate - ${certData.name} (temp-${tempTimestamp})`;
          
          const newDoc = templateDoc.makeCopy(tempName, destFolder);
          const doc = DocumentApp.openById(newDoc.getId());
          const body = doc.getBody();

          // Replace placeholders with employee name and current date
          body.replaceText('<<NAME>>', certData.name);
          body.replaceText('<<DATE>>', certData.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy"));
         
          // Save and close the document
          doc.saveAndClose();
         
          // Convert to PDF
          const pdfBlob = newDoc.getAs('application/pdf');
          pdfFile = destFolder.createFile(pdfBlob).setName(certificateName);
          pdfFileId = pdfFile.getId();
         
          // Delete the temporary Google Doc after PDF creation
          DriveApp.getFileById(newDoc.getId()).setTrashed(true);
         
          // Add the newly created file to our map to prevent duplicates within the same batch
          existingCertificatesByFolder[destFolderId].set(certificateName.toLowerCase(), pdfFileId);
         
          certificatesGenerated++;
          Logger.log(`${certData.examType} Certificate automatically created for ${certData.name} with date: ${certData.date}`);
        }
        
        // Send email if enabled and we have an email address
        if (certData.sendEmail && certData.email && pdfFileId) {
          // Create an email tracking key
          const emailKey = `EMAIL_SENT_${certData.name.replace(/\s+/g, '_')}_${certData.examType}_${certData.email}`;
          
          // Check if this email has already been sent
          if (props.getProperty(emailKey)) {
            Logger.log(`Email already sent to ${certData.email} for ${certData.examType} certificate. Skipping duplicate email.`);
          } else {
            try {
              const subject = EMAIL_SUBJECT_TEMPLATE.replace('%s', certData.examType);
              const plainBody = EMAIL_BODY_TEMPLATE.replace(/%s/g, certData.name).replace(/%s/g, certData.examType).replace(/%s/g, certData.examType);
              
              // Create HTML email with proper formatting - matching the format in generateCertificatesByType
              const htmlBody = `
                <div style="font-family: Arial, sans-serif; line-height: 1.6;">
                  <p>Dear ${certData.name},</p>
                  <p>Congratulations on successfully completing the Excel ${certData.examType} Certification Training.</p>
                  <p>We are pleased to present your official certification document, which is attached to this email. This certification recognizes your proficiency with Microsoft Excel and validates your expertise at the ${certData.examType} level.</p>
                  <p>Your achievement demonstrates both your commitment to developing valuable data analysis skills and your investment in expanding your professional capabilities. We encourage you to add this Excel certification to your professional profile and resume.</p>
                  <p>If you have any questions regarding your certification or wish to explore additional Excel training opportunities, please do not hesitate to contact us.</p>
                  <p>Best regards,</p>
                  <p>Training Certification Team</p>
                  <hr style="border: 0; border-top: 1px solid #cccccc; margin: 20px 0;">
                  
                  <!-- Email Signature -->
                  <table cellpadding="0" cellspacing="0" border="0" style="font-family: Arial, sans-serif; max-width: 500px;">
                    <tr>
                      <!-- Left column with logo -->
                      <td style="vertical-align: top; width: 150px;">
                        <img src="https://drive.google.com/uc?export=view&id=1Ato1vcuVK4PaxRDOFibaTH38OZnYHnei" alt="Aretex Logo" style="width: 150px; height: auto;">
                      </td>
                      
                      <!-- Right column with name and title -->
                      <td style="vertical-align: top; padding-left: 15px;">
                        <div style="font-size: 16px; font-weight: bold; color: #ff6600;">
                          Miki H. Burro
                        </div>
                        <div style="font-size: 12px; font-weight: bold; color: #333333; margin-top: 2px; margin-bottom: 4px;">
                          WORKFORCE EXPERIENCE - STAFF II
                        </div>
                      </td>
                    </tr>
                    
                    <!-- Tagline and contact info row -->
                    <tr>
                      <td style="font-size: 11px; color: #ff6600; font-style: italic; padding-top: 5px; white-space: nowrap;">
                        Driven by Technology. Delivered by People.
                      </td>
                      <td style="vertical-align: top; padding-left: 15px; padding-top: 5px;">
                        <div style="font-size: 12px;">
                          <a href="mailto:miki.burro@aretex.com.au" style="color: #0066cc; text-decoration: none;">miki.burro@aretex.com.au</a> | 
                          <span>+639955190265</span>
                        </div>
                      </td>
                    </tr>
                    
                    <!-- Social media row -->
                    <tr>
                      <td colspan="2" style="padding-top: 5px;">
                        <div style="background-color: #2a3698; padding: 8px; text-align: right;">
                          <a href="https://www.facebook.com" style="display: inline-block; margin-right: 5px;">
                            <img src="https://cdn-icons-png.flaticon.com/512/5968/5968764.png" alt="Facebook" style="width: 20px; height: 20px;">
                          </a>
                          <a href="https://www.linkedin.com" style="display: inline-block; margin-right: 5px;">
                            <img src="https://upload.wikimedia.org/wikipedia/commons/c/ca/LinkedIn_logo_initials.png" alt="LinkedIn" style="width: 20px; height: 20px;">
                          </a>
                          <a href="https://www.aretex.com.au" style="display: inline-block;">
                            <img src="https://cdn-icons-png.flaticon.com/512/11024/11024036.png" alt="Website" style="width: 20px; height: 20px;">
                          </a>
                        </div>
                      </td>
                    </tr>
                  </table>
                </div>
              `;
              
              // Get the PDF as a blob for attachment
              const pdfBlob = pdfFile.getBlob();

              // Send the email with the certificate attached
              GmailApp.sendEmail(
                certData.email,
                subject,
                plainBody, 
                {
                  htmlBody: htmlBody, 
                  attachments: [pdfBlob],
                  name: 'Training Certification Team'
                }
              );

              // Record that this email has been sent
              props.setProperty(emailKey, new Date().toISOString());

              emailsSent++;
              Logger.log(`Email sent to ${certData.email} with ${certData.examType} certificate for ${certData.name}`);
            } catch (emailError) {
              Logger.log(`Error sending email to ${certData.email}: ${emailError.toString()}`);
            }
          }
        }
        
        // Clean up the property regardless
        props.deleteProperty(key);
      } catch (e) {
        Logger.log(`Error processing certificate request ${key}: ${e.toString()}`);
      }
    }
  }
 
  // Clean up the trigger
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'processAutoGenerateCertificate') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
 
  if (certificatesGenerated > 0 || emailsSent > 0) {
    Logger.log(`Auto-generated ${certificatesGenerated} certificates and sent ${emailsSent} emails`);
  }
}

function generateSingleCertificate(name, examType, templateId, date, email = null, sendEmail = false) {
  // Determine the appropriate destination folder
  let destFolderId;
  if (examType === 'Basic') {
    destFolderId = "1giX-nYnriLX9IemmGpNXHiCtafProbTo";
  } else if (examType === 'Intermediate') {
    destFolderId = "171I3Ll59dNHCFxhE7wkg3GPxtfwg_fnv";
  } else if (examType === 'Advanced') {
    destFolderId = "1f0XCRnGgmFPkOVsHHilm7B8Z5er3keic";
  } else {
    Logger.log(`Invalid exam type: ${examType}`);
    return false;
  }
 
  try {
    const templateDoc = DriveApp.getFileById(templateId);
    const destFolder = DriveApp.getFolderById(destFolderId);
    const certificateName = `${examType} Certificate - ${name}.pdf`;
    let pdfFile = null;
   
    // Check if certificate already exists to avoid duplicates
    const existingFiles = destFolder.getFiles();
    let foundExisting = false;
    while (existingFiles.hasNext()) {
      const file = existingFiles.next();
      if (file.getName().trim().toLowerCase() === certificateName.trim().toLowerCase()) {
        Logger.log(`${examType} Certificate for ${name} already exists in the ${examType} folder, using existing for email`);
        pdfFile = file;
        foundExisting = true;
        break;
      }
    }
    
    if (!foundExisting) {
      // Create a unique temporary name with timestamp to prevent conflicts
      const tempTimestamp = new Date().getTime();
      const tempName = `${examType} Certificate - ${name} (temp-${tempTimestamp})`;
      
      // Make a copy of the template for the employee's certificate
      const newDoc = templateDoc.makeCopy(tempName, destFolder);
      const doc = DocumentApp.openById(newDoc.getId());
      const body = doc.getBody();

      // Replace placeholders with employee name and current date
      body.replaceText('<<NAME>>', name);
      body.replaceText('<<DATE>>', date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy"));
     
      // Save and close the document
      doc.saveAndClose();
     
      // Convert to PDF
      const pdfBlob = newDoc.getAs('application/pdf');
      pdfFile = destFolder.createFile(pdfBlob).setName(certificateName);
     
      // Delete the temporary Google Doc after PDF creation
      DriveApp.getFileById(newDoc.getId()).setTrashed(true);
     
      Logger.log(`${examType} Certificate created for ${name} with date: ${date} in the ${examType} folder`);
    }
    
    // Send email if requested and we have an email address
    if (sendEmail && email && pdfFile) {
      // Create an email tracking key
      const emailKey = `EMAIL_SENT_${name.replace(/\s+/g, '_')}_${examType}_${email}`;
      const props = PropertiesService.getScriptProperties();
      
      // Check if this email has already been sent
      if (props.getProperty(emailKey)) {
        Logger.log(`Email already sent to ${email} for ${examType} certificate. Skipping duplicate email.`);
      } else {
        try {
          const subject = EMAIL_SUBJECT_TEMPLATE.replace('%s', examType);
          const plainBody = EMAIL_BODY_TEMPLATE.replace(/%s/g, name).replace(/%s/g, examType).replace(/%s/g, examType);
          
          // Create HTML email with proper formatting - matching the format in generateCertificatesByType
          const htmlBody = `
            <div style="font-family: Arial, sans-serif; line-height: 1.6;">
              <p>Dear ${name},</p>
              <p>Congratulations on successfully completing the Excel ${examType} Certification Training.</p>
              <p>We are pleased to present your official certification document, which is attached to this email. This certification recognizes your proficiency with Microsoft Excel and validates your expertise at the ${examType} level.</p>
              <p>Your achievement demonstrates both your commitment to developing valuable data analysis skills and your investment in expanding your professional capabilities. We encourage you to add this Excel certification to your professional profile and resume.</p>
              <p>If you have any questions regarding your certification or wish to explore additional Excel training opportunities, please do not hesitate to contact us.</p>
              <p>Best regards,</p>
              <p>Training Certification Team</p>
              <hr style="border: 0; border-top: 1px solid #cccccc; margin: 20px 0;">
              
              <!-- Email Signature -->
              <table cellpadding="0" cellspacing="0" border="0" style="font-family: Arial, sans-serif; max-width: 500px;">
                <tr>
                  <!-- Left column with logo -->
                  <td style="vertical-align: top; width: 150px;">
                    <img src="https://drive.google.com/uc?export=view&id=1Ato1vcuVK4PaxRDOFibaTH38OZnYHnei" alt="Aretex Logo" style="width: 150px; height: auto;">
                  </td>
                  
                  <!-- Right column with name and title -->
                  <td style="vertical-align: top; padding-left: 15px;">
                    <div style="font-size: 16px; font-weight: bold; color: #ff6600;">
                      Miki H. Burro
                    </div>
                    <div style="font-size: 12px; font-weight: bold; color: #333333; margin-top: 2px; margin-bottom: 4px;">
                      WORKFORCE EXPERIENCE - STAFF II
                    </div>
                  </td>
                </tr>
                
                <!-- Tagline and contact info row -->
                <tr>
                  <td style="font-size: 11px; color: #ff6600; font-style: italic; padding-top: 5px; white-space: nowrap;">
                    Driven by Technology. Delivered by People.
                  </td>
                  <td style="vertical-align: top; padding-left: 15px; padding-top: 5px;">
                    <div style="font-size: 12px;">
                      <a href="mailto:miki.burro@aretex.com.au" style="color: #0066cc; text-decoration: none;">miki.burro@aretex.com.au</a> | 
                      <span>+639955190265</span>
                    </div>
                  </td>
                </tr>
                
                <!-- Social media row -->
                <tr>
                  <td colspan="2" style="padding-top: 5px;">
                    <div style="background-color: #2a3698; padding: 8px; text-align: right;">
                      <a href="https://www.facebook.com" style="display: inline-block; margin-right: 5px;">
                        <img src="https://cdn-icons-png.flaticon.com/512/5968/5968764.png" alt="Facebook" style="width: 20px; height: 20px;">
                      </a>
                      <a href="https://www.linkedin.com" style="display: inline-block; margin-right: 5px;">
                        <img src="https://upload.wikimedia.org/wikipedia/commons/c/ca/LinkedIn_logo_initials.png" alt="LinkedIn" style="width: 20px; height: 20px;">
                      </a>
                      <a href="https://www.aretex.com.au" style="display: inline-block;">
                        <img src="https://cdn-icons-png.flaticon.com/512/11024/11024036.png" alt="Website" style="width: 20px; height: 20px;">
                      </a>
                    </div>
                  </td>
                </tr>
              </table>
            </div>
          `;
          
          // Get the PDF as a blob for attachment
          const pdfBlob = pdfFile.getBlob();
          
          // Send the email with the certificate attached
          GmailApp.sendEmail(
            email,
            subject,
            plainBody,
            {
              htmlBody: htmlBody,
              attachments: [pdfBlob],
              name: 'Training Certification Team'
            }
          );
          
          // Record that this email has been sent
          props.setProperty(emailKey, new Date().toISOString());
          
          Logger.log(`Email sent to ${email} with ${examType} certificate for ${name}`);
          return true;
        } catch (emailError) {
          Logger.log(`Error sending email to ${email}: ${emailError.toString()}`);
          return false;
        }
      }
    }
   
    return true;
  } catch (e) {
    Logger.log(`Error creating ${examType} certificate for ${name}: ${e.toString()}`);
    return false;
  }
}
// Functions to generate certificates in bulk based on cell background color
function generateAllCertificates() {
  generateCertificatesByType('all');
}

function generateBasicCertificates() {
  generateCertificatesByType('Basic');
}

function generateIntermediateCertificates() {
  generateCertificatesByType('Intermediate');
}

function generateAdvancedCertificates() {
  generateCertificatesByType('Advanced');
}

function generateCertificatesByType(type) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet not found.');
    return;
  }

  const isAutoEmailEnabled = PropertiesService.getScriptProperties().getProperty('AUTO_EMAIL_ENABLED') === 'true';

  // Process each type as requested
  let generated = 0;
  let skipped = 0;
  let emailsSent = 0;

  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy");

  // Get all rows from the sheet
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const backgrounds = dataRange.getBackgrounds();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[NAME_COLUMN - 1];

    if (!name) continue;

    // Determine which certificates to generate based on type parameter
    const typesToGenerate = [];
    if (type === 'all') {
      // Check all three certificate types
      if (backgrounds[i][BASIC_SCORE_COLUMN - 1].toLowerCase() === GREEN_COLOR) {
        typesToGenerate.push({
          examType: 'Basic',
          templateId: BASIC_TEMPLATE_ID,
          destFolderId: "1giX-nYnriLX9IemmGpNXHiCtafProbTo",
          questionSheetName: 'Basic Questions',
          statusColumn: 8
        });
      }
      if (backgrounds[i][INTERMEDIATE_SCORE_COLUMN - 1].toLowerCase() === GREEN_COLOR) {
        typesToGenerate.push({
          examType: 'Intermediate',
          templateId: INTERMEDIATE_TEMPLATE_ID,
          destFolderId: "171I3Ll59dNHCFxhE7wkg3GPxtfwg_fnv",
          questionSheetName: 'Intermediate Questions',
          statusColumn: 9
        });
      }
      if (backgrounds[i][ADVANCED_SCORE_COLUMN - 1].toLowerCase() === GREEN_COLOR) {
        typesToGenerate.push({
          examType: 'Advanced',
          templateId: ADVANCED_TEMPLATE_ID,
          destFolderId: "1f0XCRnGgmFPkOVsHHilm7B8Z5er3keic",
          questionSheetName: 'Advanced Questions',
          statusColumn: 10
        });
      }
    } else {
      // Check only the specific certificate type
      let scoreColumn, templateId, destFolderId, questionSheetName, statusColumn;

      if (type === 'Basic') {
        scoreColumn = BASIC_SCORE_COLUMN;
        templateId = BASIC_TEMPLATE_ID;
        destFolderId = "1giX-nYnriLX9IemmGpNXHiCtafProbTo";
        questionSheetName = 'Basic Questions';
        statusColumn = 8;
      } else if (type === 'Intermediate') {
        scoreColumn = INTERMEDIATE_SCORE_COLUMN;
        templateId = INTERMEDIATE_TEMPLATE_ID;
        destFolderId = "171I3Ll59dNHCFxhE7wkg3GPxtfwg_fnv";
        questionSheetName = 'Intermediate Questions';
        statusColumn = 9;
      } else if (type === 'Advanced') {
        scoreColumn = ADVANCED_SCORE_COLUMN;
        templateId = ADVANCED_TEMPLATE_ID;
        destFolderId = "1f0XCRnGgmFPkOVsHHilm7B8Z5er3keic";
        questionSheetName = 'Advanced Questions';
        statusColumn = 10;
      }

      // Check if this specific certificate type should be generated
      if (backgrounds[i][scoreColumn - 1].toLowerCase() === GREEN_COLOR) {
        typesToGenerate.push({
          examType: type,
          templateId: templateId,
          destFolderId: destFolderId,
          questionSheetName: questionSheetName,
          statusColumn: statusColumn
        });
      }
    }

    // Process all certificate types that need to be generated for this person
    for (const certInfo of typesToGenerate) {
      // Fetch the email from the respective question sheet
      const questionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(certInfo.questionSheetName);
      if (!questionSheet) {
        Logger.log(`${certInfo.questionSheetName} sheet not found.`);
        continue;
      }

      const questionData = questionSheet.getDataRange().getValues();
      let email = null;

      for (let j = 1; j < questionData.length; j++) {
        if (questionData[j][2] === name) {
          email = questionData[j][3];
          break;
        }
      }

      if (!email) {
        Logger.log(`Email not found for ${name} in ${certInfo.questionSheetName} sheet.`);
        // Update the status column to indicate email not found
        sheet.getRange(i + 1, certInfo.statusColumn).setValue('Email not found');
        continue;
      }

      // Generate the certificate
      try {
        const templateDoc = DriveApp.getFileById(certInfo.templateId);
        const destFolder = DriveApp.getFolderById(certInfo.destFolderId);
        const certificateName = `${certInfo.examType} Certificate - ${name}.pdf`;
        let pdfFile = null;

        // Check if certificate already exists to avoid duplicates
        const existingFiles = destFolder.getFiles();
        let foundExisting = false;
        while (existingFiles.hasNext()) {
          const file = existingFiles.next();
          if (file.getName().trim().toLowerCase() === certificateName.trim().toLowerCase()) {
            Logger.log(`${certInfo.examType} Certificate for ${name} already exists in the ${certInfo.examType} folder, using existing for email`);
            pdfFile = file;
            foundExisting = true;
            break;
          }
        }

        if (!foundExisting) {
          // Create a unique temporary name with timestamp to prevent conflicts
          const tempTimestamp = new Date().getTime();
          const tempName = `${certInfo.examType} Certificate - ${name} (temp-${tempTimestamp})`;

          // Make a copy of the template for the employee's certificate
          const newDoc = templateDoc.makeCopy(tempName, destFolder);
          const doc = DocumentApp.openById(newDoc.getId());
          const body = doc.getBody();

          // Replace placeholders with employee name and current date
          body.replaceText('<<NAME>>', name);
          body.replaceText('<<DATE>>', date);

          // Save and close the document
          doc.saveAndClose();

          // Convert to PDF
          const pdfBlob = newDoc.getAs('application/pdf');
          pdfFile = destFolder.createFile(pdfBlob).setName(certificateName);

          // Delete the temporary Google Doc after PDF creation
          DriveApp.getFileById(newDoc.getId()).setTrashed(true);

          Logger.log(`${certInfo.examType} Certificate created for ${name} with date: ${date} in the ${certInfo.examType} folder`);
          generated++;
          // Update the status column to indicate certificate created
          sheet.getRange(i + 1, certInfo.statusColumn).setValue('Certificate created');
        } else {
          skipped++;
          // Update the status column to indicate certificate already exists
          sheet.getRange(i + 1, certInfo.statusColumn).setValue('Certificate already exists');
        }

        // Send email if requested and we have an email address
        if (isAutoEmailEnabled && email && pdfFile) {
          // Create an email tracking key
          const emailKey = `EMAIL_SENT_${name.replace(/\s+/g, '_')}_${certInfo.examType}_${email}`;
          const props = PropertiesService.getScriptProperties();

          // Check if this email has already been sent
          if (props.getProperty(emailKey)) {
            Logger.log(`Email already sent to ${email} for ${certInfo.examType} certificate. Skipping duplicate email.`);
            // Update the status column to indicate email already sent
            sheet.getRange(i + 1, certInfo.statusColumn).setValue(`Certificate already sent to ${email}`);
          } else {
            try {
              const subject = EMAIL_SUBJECT_TEMPLATE.replace('%s', certInfo.examType);
              const plainBody = EMAIL_BODY_TEMPLATE.replace(/%s/g, name).replace(/%s/g, certInfo.examType).replace(/%s/g, certInfo.examType);

              // Create HTML email with proper formatting
              const htmlBody = `
                <div style="font-family: Arial, sans-serif; line-height: 1.6;">
                  <p>Dear ${name},</p>
                  <p>Congratulations on successfully completing the Excel ${certInfo.examType} Certification Training.</p>
                  <p>We are pleased to present your official certification document, which is attached to this email. This certification recognizes your proficiency with Microsoft Excel and validates your expertise at the ${certInfo.examType} level.</p>
                  <p>Your achievement demonstrates both your commitment to developing valuable data analysis skills and your investment in expanding your professional capabilities. We encourage you to add this Excel certification to your professional profile and resume.</p>
                  <p>If you have any questions regarding your certification or wish to explore additional Excel training opportunities, please do not hesitate to contact us.</p>
                  <p>Best regards,</p>
                  <p>Training Certification Team</p>
                  <hr style="border: 0; border-top: 1px solid #cccccc; margin: 20px 0;">
                  
                  <!-- Email Signature -->
                  <table cellpadding="0" cellspacing="0" border="0" style="font-family: Arial, sans-serif; max-width: 500px;">
                    <tr>
                      <!-- Left column with logo -->
                      <td style="vertical-align: top; width: 150px;">
                        <img src="https://drive.google.com/uc?export=view&id=1Ato1vcuVK4PaxRDOFibaTH38OZnYHnei" alt="Aretex Logo" style="width: 150px; height: auto;">
                      </td>
                      
                      <!-- Right column with name and title -->
                      <td style="vertical-align: top; padding-left: 15px;">
                        <div style="font-size: 16px; font-weight: bold; color: #ff6600;">
                          Miki H. Burro
                        </div>
                        <div style="font-size: 12px; font-weight: bold; color: #333333; margin-top: 2px; margin-bottom: 4px;">
                          WORKFORCE EXPERIENCE - STAFF II
                        </div>
                      </td>
                    </tr>
                    
                    <!-- Tagline and contact info row -->
                    <tr>
                      <td style="font-size: 11px; color: #ff6600; font-style: italic; padding-top: 5px; white-space: nowrap;">
                        Driven by Technology. Delivered by People.
                      </td>
                      <td style="vertical-align: top; padding-left: 15px; padding-top: 5px;">
                        <div style="font-size: 12px;">
                          <a href="mailto:miki.burro@aretex.com.au" style="color: #0066cc; text-decoration: none;">miki.burro@aretex.com.au</a> | 
                          <span>+639955190265</span>
                        </div>
                      </td>
                    </tr>
                    
                    <!-- Social media row -->
                    <tr>
                      <td colspan="2" style="padding-top: 5px;">
                        <div style="background-color: #2a3698; padding: 8px; text-align: right;">
                          <a href="https://www.facebook.com" style="display: inline-block; margin-right: 5px;">
                            <img src="https://cdn-icons-png.flaticon.com/512/5968/5968764.png" alt="Facebook" style="width: 20px; height: 20px;">
                          </a>
                          <a href="https://www.linkedin.com" style="display: inline-block; margin-right: 5px;">
                            <img src="https://upload.wikimedia.org/wikipedia/commons/c/ca/LinkedIn_logo_initials.png" alt="LinkedIn" style="width: 20px; height: 20px;">
                          </a>
                          <a href="https://www.aretex.com.au" style="display: inline-block;">
                            <img src="https://cdn-icons-png.flaticon.com/512/11024/11024036.png" alt="Website" style="width: 20px; height: 20px;">
                          </a>
                        </div>
                      </td>
                    </tr>
                  </table>
                </div>
              `;

              // Get the PDF as a blob for attachment
              const pdfBlob = pdfFile.getBlob();

              // Send the email with the certificate attached
              GmailApp.sendEmail(
                email,
                subject,
                plainBody,
                {
                  htmlBody: htmlBody,
                  attachments: [pdfBlob],
                  name: 'Training Certification Team'
                }
              );

              // Record that this email has been sent
              props.setProperty(emailKey, new Date().toISOString());

              // Update the status column to indicate email sent
              sheet.getRange(i + 1, certInfo.statusColumn).setValue(`Certificate already sent to ${email}`);

              Logger.log(`Email sent to ${email} with ${certInfo.examType} certificate for ${name}`);
              emailsSent++;
            } catch (emailError) {
              Logger.log(`Error sending email to ${email}: ${emailError.toString()}`);
              // Update the status column to indicate email error
              sheet.getRange(i + 1, certInfo.statusColumn).setValue(`Error sending email: ${emailError.message}`);
            }
          }
        } else if (!isAutoEmailEnabled && email && pdfFile) {
          // Auto-email not enabled, update status only for new certificates
          const emailKey = `EMAIL_SENT_${name.replace(/\s+/g, '_')}_${certInfo.examType}_${email}`;
          const props = PropertiesService.getScriptProperties();

          // Check if the email has already been sent
          if (!props.getProperty(emailKey)) {
            const currentStatus = sheet.getRange(i + 1, certInfo.statusColumn).getValue();
            // Only update the status if it is not already "Certificate already sent"
            if (!currentStatus.includes("Certificate already sent")) {
              sheet.getRange(i + 1, certInfo.statusColumn).setValue('Certificate is available (Email not sent - auto-email disabled)');
            }
          } else {
            // If the email was already sent, keep the status as "Certificate already sent"
            const currentStatus = sheet.getRange(i + 1, certInfo.statusColumn).getValue();
            if (!currentStatus.includes("Certificate already sent")) {
              sheet.getRange(i + 1, certInfo.statusColumn).setValue(`Certificate already sent to ${email}`);
            }
          }
        }
      } catch (e) {
        Logger.log(`Error creating ${certInfo.examType} certificate for ${name}: ${e.toString()}`);
        // Update the status column to indicate generation error
        sheet.getRange(i + 1, certInfo.statusColumn).setValue(`Error: ${e.message}`);
      }
    }
  }

  SpreadsheetApp.getUi().alert(`Certificate generation complete!\nGenerated: ${generated}\nSkipped: ${skipped}\nEmails Sent: ${emailsSent}`);
}

// Add menu options
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Certificates')
    .addItem('Generate All Types of Certificates', 'generateAllCertificates')
    .addSeparator()
    .addItem('Generate Basic Exam Certificates', 'generateBasicCertificates')
    .addItem('Generate Intermediate Exam Certificates', 'generateIntermediateCertificates')
    .addItem('Generate Advanced Exam Certificates', 'generateAdvancedCertificates')
    .addSeparator()
    .addItem('Enable/Disable Automatic Certificate Generation', 'toggleAutoCertificateGeneration')
    .addItem('Enable/Disable Automatic Email Sending', 'toggleAutoEmailSending')
    .addToUi();
}

function toggleAutoCertificateGeneration() {
  const ui = SpreadsheetApp.getUi();
  const scriptProps = PropertiesService.getScriptProperties();
 
  // Check current auto-generation status
  const isEnabled = scriptProps.getProperty('AUTO_CERT_GENERATION_ENABLED') === 'true';
 
  // Create the dialog content
  let message = isEnabled ?
    "Automatic certificate generation is currently ENABLED.\n\nWhen a score cell is highlighted GREEN, certificates will be automatically generated.\n\nWould you like to disable this feature?" :
    "Automatic certificate generation is currently DISABLED.\n\nWould you like to enable automatic certificate generation when score cells are highlighted GREEN?";
 
  // Show confirmation dialog with correct parameters
  const response = ui.alert('Certificate Auto-Generation Settings', message, ui.ButtonSet.YES_NO);
 
  // Process the user's choice
  if (response === ui.Button.YES) {
    // Toggle the setting
    scriptProps.setProperty('AUTO_CERT_GENERATION_ENABLED', (!isEnabled).toString());
   
    // Confirm the change with correct parameters - include ButtonSet
    const newStatus = !isEnabled ? 'ENABLED' : 'DISABLED';
    ui.alert('Settings Updated', `Automatic certificate generation is now ${newStatus}.`, ui.ButtonSet.OK);
  }
}

function toggleAutoEmailSending() {
  const ui = SpreadsheetApp.getUi();
  const scriptProps = PropertiesService.getScriptProperties();
  
  // Check current auto-email status
  const isEnabled = scriptProps.getProperty('AUTO_EMAIL_ENABLED') === 'true';
  
  // Create the dialog content
  let message = isEnabled ?
    "Automatic email sending is currently ENABLED.\n\nWhen certificates are generated, they will be automatically emailed to employees.\n\nWould you like to disable this feature?" :
    "Automatic email sending is currently DISABLED.\n\nWould you like to enable automatic email sending when certificates are generated?";
    
  // Show confirmation dialog
  const response = ui.alert('Email Auto-Sending Settings', message, ui.ButtonSet.YES_NO);
  
  // Process the user's choice
  if (response === ui.Button.YES) {
    // Toggle the setting
    scriptProps.setProperty('AUTO_EMAIL_ENABLED', (!isEnabled).toString());
    
    // Confirm the change
    const newStatus = !isEnabled ? 'ENABLED' : 'DISABLED';
    ui.alert('Settings Updated', `Automatic email sending is now ${newStatus}.`, ui.ButtonSet.OK);
    
    // Additional check if they're enabling emails but certificate generation is disabled
    if (!isEnabled && scriptProps.getProperty('AUTO_CERT_GENERATION_ENABLED') !== 'true') {
      ui.alert('Note', 'You have enabled automatic email sending, but automatic certificate generation is still disabled. If you want certificates to be automatically generated and emailed when scores are highlighted green, please also enable automatic certificate generation.', ui.ButtonSet.OK);
    }
  }
}