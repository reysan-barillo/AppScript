
function testCertificateExpirationNotices() {
    const NOTIFICATION_DAYS = 30; // Days before expiration to send notification
    
    // For testing, we can use a custom testing property to track sent notifications
    const TEST_PREFIX = "TEST_EXPIRATION_NOTICE_";
    
    // Folder IDs for each certificate type
    const FOLDER_IDS = {
      "Basic": "1giX-nYnriLX9IemmGpNXHiCtafProbTo",
      "Intermediate": "171I3Ll59dNHCFxhE7wkg3GPxtfwg_fnv", 
      "Advanced": "1f0XCRnGgmFPkOVsHHilm7B8Z5er3keic"
    };
    
    // Email template
    const EMAIL_SUBJECT = "[TEST] Your Excel %s Certification is About to Expire";
    const EMAIL_BODY = `Dear %s,
  
  This is a TEST notification.
  
  Your Excel %s Certification will expire on %s (in approximately %s days).
  
  To maintain your certified status, please consider scheduling a recertification exam at your earliest convenience.
  
  If you have any questions about the recertification process, please contact us.
  
  Best regards,
  Training Certification Team`;
  
    const HTML_EMAIL_BODY = `
      <div style="font-family: Arial, sans-serif; line-height: 1.6;">
        <p><strong>THIS IS A TEST NOTIFICATION</strong></p>
        <p>Dear %s,</p>
        <p>Your Excel <strong>%s Certification</strong> will expire on <strong>%s</strong> (in approximately <strong>%s days</strong>).</p>
        <p>To maintain your certified status, please consider scheduling a recertification exam at your earliest convenience.</p>
        <p>If you have any questions about the recertification process, please contact us.</p>
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
          </tr>
          <!-- Tagline and contact info row -->
          <tr>
            <td style="font-size: 11px; color: #ff6600; font-style: italic; padding-top: 5px; white-space: nowrap;">
              Driven by Technology. Delivered by People.
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
    
    // Get the current date for comparison
    const currentDate = new Date();
    
    // Track the counts
    let noticesSent = 0;
    let notificationsToSend = [];
    
    // For testing purposes, we'll create a recipient list
    // Either set up test emails or use your own email for testing
    const testRecipients = {
      "Test User": "your-test-email@example.com"  // Replace with your email for testing
    };
    
    // Process each certificate type
    for (const [examType, folderId] of Object.entries(FOLDER_IDS)) {
      try {
        // For testing, we'll use a simulated expiration date for a specific certificate type
        // This way we can test without actually modifying any real certificates
        // Here, we're simulating that all certificates will expire in 15 days (within our notification window)
        const simulatedExpirationDate = new Date(currentDate);
        simulatedExpirationDate.setDate(simulatedExpirationDate.getDate() + 15); // 15 days from now
        
        const daysUntilExpiration = 15; // Simulated days until expiration
        
        // For testing, we'll send a notification for each test recipient
        for (const [employeeName, employeeEmail] of Object.entries(testRecipients)) {
          // Format dates for readability
          const formattedExpirationDate = Utilities.formatDate(simulatedExpirationDate, Session.getScriptTimeZone(), "MMMM d, yyyy");
          
          // Create a test notification key
          const notificationKey = `${TEST_PREFIX}${employeeName.replace(/\s+/g, '_')}_${examType}_${formattedExpirationDate}`;
          
          notificationsToSend.push({
            name: employeeName,
            email: employeeEmail,
            examType: examType,
            expirationDate: formattedExpirationDate,
            daysRemaining: daysUntilExpiration,
            notificationKey: notificationKey
          });
        }
      } catch (error) {
        Logger.log(`Error processing ${examType} certificates: ${error.toString()}`);
      }
    }
    
    // Log and display notification information before sending
    Logger.log(`Prepared ${notificationsToSend.length} test notifications:`);
    for (const notification of notificationsToSend) {
      Logger.log(`- ${notification.name} (${notification.email}): ${notification.examType} expires on ${notification.expirationDate}`);
    }
  
    // Optional: Add a UI confirmation before sending test emails
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Send Test Notifications', 
      `Ready to send ${notificationsToSend.length} test notifications. Continue?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.NO) {
      return `Test canceled. ${notificationsToSend.length} notifications were prepared but not sent.`;
    }
    
    // Now send all the test notifications
    for (const notification of notificationsToSend) {
      try {
        // Format email content
        const subject = EMAIL_SUBJECT.replace('%s', notification.examType);
        const plainBody = EMAIL_BODY
          .replace('%s', notification.name)
          .replace('%s', notification.examType)
          .replace('%s', notification.expirationDate)
          .replace('%s', notification.daysRemaining);
        
        const htmlBody = HTML_EMAIL_BODY
          .replace('%s', notification.name)
          .replace('%s', notification.examType)
          .replace('%s', notification.expirationDate)
          .replace('%s', notification.daysRemaining);
        
        // Send email
        GmailApp.sendEmail(
          notification.email,
          subject,
          plainBody,
          {
            htmlBody: htmlBody,
            name: 'Training Certification Team [TEST]'
          }
        );
        
        // For test purposes, we'll track these separately from real notifications
        PropertiesService.getScriptProperties().setProperty(notification.notificationKey, new Date().toISOString());
        
        noticesSent++;
        Logger.log(`TEST expiration notice sent to ${notification.email} for ${notification.examType} certificate (expires on ${notification.expirationDate})`);
      } catch (emailError) {
        Logger.log(`Error sending test expiration notice to ${notification.email}: ${emailError.toString()}`);
      }
    }
    
    Logger.log(`Test certificate expiration notice process complete. Sent ${noticesSent} notifications.`);
    
    // Return summary
    return `Test completed. Sent ${noticesSent} test notifications for certificates simulated to expire in 15 days.`;
  }
  
  
  function testWithRealCertificates() {
    const NOTIFICATION_DAYS = 30; // Days before expiration to send notification
    const TEST_EMAIL = "your-test-email@example.com"; // Replace with your email for testing
    
    // Folder IDs for each certificate type
    const FOLDER_IDS = {
      "Basic": "1giX-nYnriLX9IemmGpNXHiCtafProbTo",
      "Intermediate": "171I3Ll59dNHCFxhE7wkg3GPxtfwg_fnv", 
      "Advanced": "1f0XCRnGgmFPkOVsHHilm7B8Z5er3keic"
    };
    
    // Email template with TEST indicator
    const EMAIL_SUBJECT = "[TEST] Your Excel %s Certification is About to Expire";
    const EMAIL_BODY = `Dear %s,
  
  THIS IS A TEST NOTIFICATION using real certificate data.
  
  Your Excel %s Certification would normally expire on %s, but for this test we're simulating it will expire in %s days.
  
  To maintain your certified status, please consider scheduling a recertification exam at your earliest convenience.
  
  If you have any questions about the recertification process, please contact us.
  
  Best regards,
  Training Certification Team`;
  
    const HTML_EMAIL_BODY = `
      <div style="font-family: Arial, sans-serif; line-height: 1.6;">
        <p><strong>THIS IS A TEST NOTIFICATION</strong> using real certificate data.</p>
        <p>Dear %s,</p>
        <p>Your Excel <strong>%s Certification</strong> would normally expire on <strong>%s</strong>, but for this test we're simulating it will expire in <strong>%s days</strong>.</p>
        <p>To maintain your certified status, please consider scheduling a recertification exam at your earliest convenience.</p>
        <p>If you have any questions about the recertification process, please contact us.</p>
        <p>Best regards,</p>
        <p>Training Certification Team</p>
      </div>
    `;
    
    // Get the current date for comparison
    const currentDate = new Date();
    
    // Track the counts
    let noticesSent = 0;
    let notificationsToSend = [];
    let certificatesFound = 0;
    
    // Process each certificate type
    for (const [examType, folderId] of Object.entries(FOLDER_IDS)) {
      try {
        const folder = DriveApp.getFolderById(folderId);
        const files = folder.getFiles();
        
        while (files.hasNext()) {
          const file = files.next();
          certificatesFound++;
          
          // Only process PDF files (certificates)
          if (file.getMimeType() === "application/pdf" && file.getName().includes("Certificate")) {
            const creationDate = file.getDateCreated();
            
            // Calculate actual expiration date (1 year after creation)
            const actualExpirationDate = new Date(creationDate);
            actualExpirationDate.setFullYear(actualExpirationDate.getFullYear() + 1);
            
            // For testing, we'll force the expiration date to be 15 days from now
            const testExpirationDate = new Date(currentDate);
            testExpirationDate.setDate(testExpirationDate.getDate() + 15);
            
            const daysUntilExpiration = 15; // Fixed for testing
            
            // Extract employee name from the certificate filename
            // Format is typically "{ExamType} Certificate - {Name}.pdf"
            const fileName = file.getName();
            const nameMatch = fileName.match(/Certificate - (.+?)\.pdf$/);
            
            if (nameMatch && nameMatch[1]) {
              const employeeName = nameMatch[1];
              
              // Format dates for readability
              const formattedActualExpiration = Utilities.formatDate(actualExpirationDate, Session.getScriptTimeZone(), "MMMM d, yyyy");
              const formattedTestExpiration = Utilities.formatDate(testExpirationDate, Session.getScriptTimeZone(), "MMMM d, yyyy");
              
              // Create a test notification key
              const notificationKey = `TEST_FORCED_EXPIRATION_${employeeName.replace(/\s+/g, '_')}_${examType}_${formattedTestExpiration}`;
              
              notificationsToSend.push({
                name: employeeName,
                email: TEST_EMAIL, // For testing, all emails go to test email
                examType: examType,
                actualExpirationDate: formattedActualExpiration,
                testExpirationDate: formattedTestExpiration,
                daysRemaining: daysUntilExpiration,
                notificationKey: notificationKey
              });
            }
          }
        }
      } catch (error) {
        Logger.log(`Error processing ${examType} certificates: ${error.toString()}`);
      }
    }
    
    // Log details before sending
    Logger.log(`Found ${certificatesFound} certificates. Prepared ${notificationsToSend.length} test notifications.`);
    
    // Optional: Add a UI confirmation before sending test emails
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Send Test Notifications', 
      `Found ${certificatesFound} certificates. Ready to send ${notificationsToSend.length} test notifications to ${TEST_EMAIL}. Continue?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.NO) {
      return `Test canceled. ${notificationsToSend.length} notifications were prepared but not sent.`;
    }
    
    // Limit the number of test emails to avoid overwhelming the inbox
    const MAX_TEST_EMAILS = 5;
    if (notificationsToSend.length > MAX_TEST_EMAILS) {
      notificationsToSend = notificationsToSend.slice(0, MAX_TEST_EMAILS);
      Logger.log(`Limiting test to ${MAX_TEST_EMAILS} notifications to avoid inbox flooding.`);
    }
    
    // Now send the test notifications
    for (const notification of notificationsToSend) {
      try {
        // Format email content
        const subject = EMAIL_SUBJECT.replace('%s', notification.examType);
        const plainBody = EMAIL_BODY
          .replace('%s', notification.name)
          .replace('%s', notification.examType)
          .replace('%s', notification.actualExpirationDate)
          .replace('%s', notification.daysRemaining);
        
        const htmlBody = HTML_EMAIL_BODY
          .replace('%s', notification.name)
          .replace('%s', notification.examType)
          .replace('%s', notification.actualExpirationDate)
          .replace('%s', notification.daysRemaining);
        
        // Send email
        GmailApp.sendEmail(
          notification.email,
          subject,
          plainBody,
          {
            htmlBody: htmlBody,
            name: 'Training Certification Team [TEST]'
          }
        );
        
        noticesSent++;
        Logger.log(`TEST notice sent to ${notification.email} for ${notification.name}'s ${notification.examType} certificate`);
      } catch (emailError) {
        Logger.log(`Error sending test notice to ${notification.email}: ${emailError.toString()}`);
      }
    }
    
    Logger.log(`Test with real certificates complete. Sent ${noticesSent} test notifications.`);
    
    // Return summary
    return `Test completed. Found ${certificatesFound} certificates and sent ${noticesSent} test notifications to ${TEST_EMAIL}.`;
  }
  
  function clearTestNotificationTracking() {
    const props = PropertiesService.getScriptProperties();
    const allProps = props.getProperties();
    
    let testPropsCount = 0;
    
    for (const key in allProps) {
      if (key.startsWith('TEST_')) {
        props.deleteProperty(key);
        testPropsCount++;
      }
    }
    
    Logger.log(`Cleared ${testPropsCount} test notification tracking records.`);
    return `Cleared ${testPropsCount} test notification tracking records.`;
  }