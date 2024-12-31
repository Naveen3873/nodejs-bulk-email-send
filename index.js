const nodemailer = require('nodemailer');
const Excel = require('exceljs');

// Create a transporter object using Gmail's SMTP service
let transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: 'your-email',  // Your email address
    pass: 'app password'        // Your Gmail App Password (NOT your regular Gmail password)
  }
});

// List of recipient emails with their names for personalization
let emails = [
  { email: 'naveenbe3873@gmail.com', name: 'Naveen' }
];

// Initialize the Excel workbook and sheet
const workbook = new Excel.Workbook();
const worksheet = workbook.addWorksheet('Sent Emails');

// Define columns for the Excel sheet
worksheet.columns = [
  { header: 'Email Address', key: 'email', width: 30 },
  { header: 'Status', key: 'status', width: 15 },
  { header: 'Date', key: 'date', width: 25 },
  { header: 'Error (if any)', key: 'error', width: 50 }
];

// Function to send emails and log results
async function sendEmails() {
  // Loop through each recipient email
  for (let i = 0; i < emails.length; i++) {
    const { email, name } = emails[i];
    // HTML email content with inline CSS
    // you can edit your content in the HTML
    const htmlContent = `
      <!DOCTYPE html>
      <html lang="en">
      <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Demo Class Invitation</title>
          <style>
              body {
                  font-family: 'Arial', sans-serif;
                  margin: 0;
                  padding: 0;
                  background-color: #f4f4f9;
                  color: #333;
              }
              .email-container {
                  width: 100%;
                  max-width: 600px;
                  margin: 0 auto;
                  background-color: #ffffff;
                  padding: 20px;
                  border-radius: 10px;
                  box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
              }
              .header {
                  text-align: center;
                  background-color: #4CAF50;
                  padding: 20px;
                  border-radius: 10px;
                  color: #fff;
              }
              .header h1 {
                  margin: 0;
                  font-size: 24px;
              }
              .body-content {
                  padding: 20px;
                  text-align: center;
              }
              .body-content h2 {
                  font-size: 22px;
                  margin-bottom: 10px;
              }
              .body-content p {
                  font-size: 16px;
                  margin-bottom: 20px;
              }
              .cta-button {
                  background-color: #4CAF50;
                  color: white;
                  padding: 15px 30px;
                  font-size: 16px;
                  border: none;
                  border-radius: 5px;
                  text-decoration: none;
                  display: inline-block;
                  margin-top: 20px;
              }
              .cta-button:hover {
                  background-color: #45a049;
              }
              .footer {
                  text-align: center;
                  margin-top: 30px;
                  font-size: 14px;
                  color: #777;
              }
              .footer a {
                  color: #4CAF50;
                  text-decoration: none;
              }
              .footer a:hover {
                  text-decoration: underline;
              }
          </style>
      </head>
      <body>
          <div class="email-container">
              <!-- Header -->
              <div class="header">
                  <h1>You're Invited to a Google Meeting</h1>
              </div>
      
              <!-- Email Body -->
              <div class="body-content">
                  <h2>Join us for a Demo Class!</h2>
                  <p>Dear ${name},</p>
                  <p>We are excited to invite you to a demo class where we will be introducing you to our services and answering any questions you might have. This session will be hosted on Google Meet.</p>
                  <p><strong>Date and Time:</strong> January 1, 2025 - 10:00 AM</p>
                  <p>Click the button below to attend the meeting:</p>
                  <a href="https://meet.google.com/example-link-123" class="cta-button">Join Google Meet</a>
              </div>
      
              <!-- Footer -->
              <div class="footer">
                  <p>If you have any questions, feel free to <a href="mailto:support@example.com">contact us</a>.</p>
                  <p>Looking forward to seeing you at the demo class!</p>
              </div>
          </div>
      </body>
      </html>
    `;

    const mailOptions = {
      from: '"Naveen K" <naveened21@gmail.com>', // Sender address
      to: email, // Recipient address
      subject: 'Test Email from Node.js', // Subject line
      html: htmlContent // HTML email body with CSS
    };

    try {
      // Send the email
      const info = await transporter.sendMail(mailOptions);
      console.log(`Email sent successfully to: ${email}`);

      // Log successful email in Excel sheet
      worksheet.addRow({
        email: email,
        status: 'Success',
        date: new Date().toLocaleString(),
        error: ''
      });
    } catch (error) {
      console.log(`Error sending email to ${email}:`, error);

      // Log failed email attempt in Excel sheet
      worksheet.addRow({
        email: email,
        status: 'Failed',
        date: new Date().toLocaleString(),
        error: error.message || 'Unknown error'
      });
    }
  }

  // Save the Excel file after processing all emails
  const fileName = `SentEmails_${new Date().toLocaleString().replace(/[^\w\s]/gi, '_')}.xlsx`;
  await workbook.xlsx.writeFile(fileName);
  console.log(`Excel file "${fileName}" saved!`);
}

// Execute the function to send emails
sendEmails();
