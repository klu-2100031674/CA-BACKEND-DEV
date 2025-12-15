const nodemailer = require('nodemailer');

const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASSWORD
  }
});

// Helper function to check if email is a test email
const isTestEmail = (email) => {
  if (!email) return false;

  const testPatterns = [
    '@test.com',
    '@example.com',
    '@mailinator.com',
    '@10minutemail.com',
    '@temp-mail.org',
    '@guerrillamail.com',
    '@maildrop.cc'
  ];

  const lowerEmail = email.toLowerCase();
  return testPatterns.some(pattern => lowerEmail.endsWith(pattern));
};

const sendVerificationEmail = async (email, token) => {
  // Skip sending email for test accounts
  if (isTestEmail(email)) {
    console.log(`Skipping verification email for test account: ${email}`);
    return;
  }
  
  const frontendUrl = process.env.FRONTEND_URL || 'http://localhost:5173';
  const verificationUrl = `${frontendUrl}/verify-email?token=${token}`;
  
  const mailOptions = {
    from: process.env.EMAIL_USER,
    to: email,
    subject: 'Email Verification',
    html: `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2>Email Verification</h2>
        <p>Thank you for registering! Please click the button below to verify your email address:</p>
        <a href="${verificationUrl}" style="background-color: #007bff; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block;">Verify Email</a>
        <p>Or copy and paste this link in your browser:</p>
        <p>${verificationUrl}</p>
        <p>This link will expire in 1 hour.</p>
      </div>
    `
  };
  
  await transporter.sendMail(mailOptions);
};

const sendWelcomeEmail = async (email, name) => {
  // Skip sending email for test accounts
  if (isTestEmail(email)) {
    console.log(`Skipping welcome email for test account: ${email}`);
    return;
  }
  
  const mailOptions = {
    from: process.env.EMAIL_USER,
    to: email,
    subject: 'Welcome to Our Platform',
    html: `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2>Welcome ${name}!</h2>
        <p>Your email has been verified successfully. You can now access all features of our platform.</p>
        <p>Happy exploring!</p>
      </div>
    `
  };
  
  await transporter.sendMail(mailOptions);
};

const sendOrderConfirmation = async (email, name, order) => {
  // Skip sending email for test accounts
  if (isTestEmail(email)) {
    console.log(`Skipping order confirmation email for test account: ${email}`);
    return;
  }
  
  const mailOptions = {
    from: process.env.EMAIL_USER,
    to: email,
    subject: 'Order Confirmation',
    html: `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2>Order Confirmed!</h2>
        <p>Hi ${name},</p>
        <p>Your order has been confirmed:</p>
        <ul>
          <li>Pack Type: ${order.pack_type}</li>
          <li>Credits: ${order.credits}</li>
          <li>Amount: ₹${order.amount_paid}</li>
          <li>Order ID: ${order._id}</li>
        </ul>
        <p>Your credits have been added to your wallet.</p>
      </div>
    `
  };
  
  await transporter.sendMail(mailOptions);
};

const sendPasswordResetOTP = async (email, otp) => {
  // Skip sending email for test accounts but log the OTP for testing
  if (isTestEmail(email)) {
    console.log(`Skipping OTP email for test account: ${email}`);
    console.log(`Test OTP for ${email}: ${otp}`);
    return;
  }
  
  const mailOptions = {
    from: process.env.EMAIL_USER,
    to: email,
    subject: 'Password Reset OTP',
    html: `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2>Password Reset Request</h2>
        <p>You have requested to reset your password. Use the OTP below to verify your identity:</p>
        <div style="background-color: #f4f4f4; padding: 20px; text-align: center; margin: 20px 0;">
          <h1 style="color: #007bff; letter-spacing: 5px; margin: 0;">${otp}</h1>
        </div>
        <p>This OTP is valid for <strong>10 minutes</strong>.</p>
        <p>If you did not request this password reset, please ignore this email or contact support.</p>
        <hr style="border: none; border-top: 1px solid #eee; margin: 20px 0;">
        <p style="color: #666; font-size: 12px;">This is an automated message. Please do not reply.</p>
      </div>
    `
  };
  
  await transporter.sendMail(mailOptions);
};

const sendReportGeneratedNotification = async (email, name, reportDetails) => {
  // Skip sending email for test accounts
  if (isTestEmail(email)) {
    console.log(`Skipping report notification email for test account: ${email}`);
    return;
  }
  
  const mailOptions = {
    from: process.env.EMAIL_USER,
    to: email,
    subject: 'Your Report is Ready!',
    html: `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2>Report Generated Successfully!</h2>
        <p>Hi ${name},</p>
        <p>Your report has been generated and is ready for download.</p>
        <div style="background-color: #f4f4f4; padding: 15px; margin: 20px 0; border-radius: 5px;">
          <p><strong>Report Type:</strong> ${reportDetails.templateId || 'Financial Report'}</p>
          <p><strong>Generated On:</strong> ${new Date().toLocaleDateString('en-IN')}</p>
          <p><strong>Report ID:</strong> ${reportDetails.reportId}</p>
        </div>
        <p>You can access your report from your dashboard.</p>
        <a href="${process.env.FRONTEND_URL || 'http://localhost:5173'}/reports" 
           style="background-color: #007bff; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block; margin-top: 10px;">
          View Reports
        </a>
        <hr style="border: none; border-top: 1px solid #eee; margin: 20px 0;">
        <p style="color: #666; font-size: 12px;">Thank you for using our service!</p>
      </div>
    `
  };
  
  await transporter.sendMail(mailOptions);
};

const sendWithdrawalStatusNotification = async (email, name, withdrawalDetails) => {
  // Skip sending email for test accounts
  if (isTestEmail(email)) {
    console.log(`Skipping withdrawal notification email for test account: ${email}`);
    return;
  }
  
  const statusMessages = {
    'approved': 'Your withdrawal request has been approved and is being processed.',
    'rejected': 'Your withdrawal request has been rejected.',
    'completed': 'Your withdrawal has been completed successfully!',
    'processing': 'Your withdrawal request is being processed.'
  };
  
  const mailOptions = {
    from: process.env.EMAIL_USER,
    to: email,
    subject: `Withdrawal Request ${withdrawalDetails.status.charAt(0).toUpperCase() + withdrawalDetails.status.slice(1)}`,
    html: `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2>Withdrawal Request Update</h2>
        <p>Hi ${name},</p>
        <p>${statusMessages[withdrawalDetails.status] || 'Your withdrawal request status has been updated.'}</p>
        <div style="background-color: #f4f4f4; padding: 15px; margin: 20px 0; border-radius: 5px;">
          <p><strong>Amount:</strong> ₹${withdrawalDetails.amount}</p>
          <p><strong>Status:</strong> ${withdrawalDetails.status.toUpperCase()}</p>
          ${withdrawalDetails.transaction_id ? `<p><strong>Transaction ID:</strong> ${withdrawalDetails.transaction_id}</p>` : ''}
          ${withdrawalDetails.admin_remarks ? `<p><strong>Remarks:</strong> ${withdrawalDetails.admin_remarks}</p>` : ''}
        </div>
        <hr style="border: none; border-top: 1px solid #eee; margin: 20px 0;">
        <p style="color: #666; font-size: 12px;">This is an automated message. Please do not reply.</p>
      </div>
    `
  };
  
  await transporter.sendMail(mailOptions);
};

/**
 * Send report approval notification email
 */
const sendReportApprovalEmail = async (email, name, reportDetails) => {
  // Skip sending email for test accounts
  if (isTestEmail(email)) {
    console.log(`Skipping report approval email for test account: ${email}`);
    return;
  }
  
  const frontendUrl = process.env.FRONTEND_URL || 'http://localhost:5173';
  
  const mailOptions = {
    from: process.env.EMAIL_USER,
    to: email,
    subject: '✅ Your Report Has Been Approved!',
    html: `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px;">
        <div style="text-align: center; margin-bottom: 30px;">
          <div style="background: linear-gradient(135deg, #10B981, #059669); width: 60px; height: 60px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center;">
            <span style="font-size: 30px;">✓</span>
          </div>
        </div>
        
        <h2 style="color: #10B981; text-align: center;">Report Approved!</h2>
        
        <p>Hi ${name},</p>
        <p>Great news! Your report has been reviewed and approved by our team.</p>
        
        <div style="background-color: #F0FDF4; border: 1px solid #BBF7D0; padding: 20px; margin: 20px 0; border-radius: 8px;">
          <p style="margin: 5px 0;"><strong>Report:</strong> ${reportDetails.title || 'Financial Report'}</p>
          <p style="margin: 5px 0;"><strong>Report Type:</strong> ${reportDetails.report_type || reportDetails.templateId}</p>
          <p style="margin: 5px 0;"><strong>Approved On:</strong> ${new Date().toLocaleDateString('en-IN', { dateStyle: 'long' })}</p>
          ${reportDetails.validation_notes ? `<p style="margin: 5px 0;"><strong>Notes:</strong> ${reportDetails.validation_notes}</p>` : ''}
        </div>
        
        <p>You can now download your report files (Excel & PDF) from your dashboard.</p>
        
        <div style="text-align: center; margin: 30px 0;">
          <a href="${frontendUrl}/reports" 
             style="background: linear-gradient(135deg, #8B5CF6, #7C3AED); color: white; padding: 14px 28px; text-decoration: none; border-radius: 8px; display: inline-block; font-weight: 600;">
            View My Reports
          </a>
        </div>
        
        <hr style="border: none; border-top: 1px solid #E5E7EB; margin: 30px 0;">
        <p style="color: #6B7280; font-size: 12px; text-align: center;">
          Thank you for using CA Excel Report Generation Service!<br>
          This is an automated message. Please do not reply.
        </p>
      </div>
    `
  };
  
  await transporter.sendMail(mailOptions);
};

/**
 * Send report rejection notification email
 */
const sendReportRejectionEmail = async (email, name, reportDetails) => {
  // Skip sending email for test accounts
  if (isTestEmail(email)) {
    console.log(`Skipping report rejection email for test account: ${email}`);
    return;
  }
  
  const frontendUrl = process.env.FRONTEND_URL || 'http://localhost:5173';
  
  const mailOptions = {
    from: process.env.EMAIL_USER,
    to: email,
    subject: '❌ Your Report Needs Revision',
    html: `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px;">
        <div style="text-align: center; margin-bottom: 30px;">
          <div style="background: linear-gradient(135deg, #EF4444, #DC2626); width: 60px; height: 60px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center;">
            <span style="font-size: 30px; color: white;">✕</span>
          </div>
        </div>
        
        <h2 style="color: #EF4444; text-align: center;">Report Needs Revision</h2>
        
        <p>Hi ${name},</p>
        <p>Unfortunately, your report could not be approved in its current form. Please review the feedback below.</p>
        
        <div style="background-color: #FEF2F2; border: 1px solid #FECACA; padding: 20px; margin: 20px 0; border-radius: 8px;">
          <p style="margin: 5px 0;"><strong>Report:</strong> ${reportDetails.title || 'Financial Report'}</p>
          <p style="margin: 5px 0;"><strong>Report Type:</strong> ${reportDetails.report_type || reportDetails.templateId}</p>
          <p style="margin: 5px 0;"><strong>Reviewed On:</strong> ${new Date().toLocaleDateString('en-IN', { dateStyle: 'long' })}</p>
        </div>
        
        <div style="background-color: #FEF3C7; border-left: 4px solid #F59E0B; padding: 15px; margin: 20px 0;">
          <p style="margin: 0; font-weight: 600; color: #92400E;">Reason for Rejection:</p>
          <p style="margin: 10px 0 0 0; color: #78350F;">${reportDetails.rejection_reason || 'Please contact support for more details.'}</p>
        </div>
        
        <p>Please make the necessary corrections and resubmit your report.</p>
        
        <div style="text-align: center; margin: 30px 0;">
          <a href="${frontendUrl}/reports" 
             style="background: linear-gradient(135deg, #8B5CF6, #7C3AED); color: white; padding: 14px 28px; text-decoration: none; border-radius: 8px; display: inline-block; font-weight: 600;">
            View My Reports
          </a>
        </div>
        
        <hr style="border: none; border-top: 1px solid #E5E7EB; margin: 30px 0;">
        <p style="color: #6B7280; font-size: 12px; text-align: center;">
          If you have questions, please contact our support team.<br>
          This is an automated message. Please do not reply.
        </p>
      </div>
    `
  };
  
  await transporter.sendMail(mailOptions);
};

/**
 * Send report with PDF attachment
 */
const sendReportWithAttachment = async (email, name, reportDetails, pdfBuffer = null) => {
  // Skip sending email for test accounts
  if (isTestEmail(email)) {
    console.log(`Skipping report attachment email for test account: ${email}`);
    return;
  }
  
  const frontendUrl = process.env.FRONTEND_URL || 'http://localhost:5173';
  
  const mailOptions = {
    from: process.env.EMAIL_USER,
    to: email,
    subject: `📊 Your Approved Report: ${reportDetails.title}`,
    html: `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px;">
        <h2 style="color: #7C3AED;">Your Report is Ready!</h2>
        
        <p>Hi ${name},</p>
        <p>Please find your approved report attached to this email.</p>
        
        <div style="background-color: #F3F4F6; padding: 20px; margin: 20px 0; border-radius: 8px;">
          <p style="margin: 5px 0;"><strong>Report:</strong> ${reportDetails.title}</p>
          <p style="margin: 5px 0;"><strong>Type:</strong> ${reportDetails.report_type || reportDetails.templateId}</p>
          <p style="margin: 5px 0;"><strong>Generated:</strong> ${new Date(reportDetails.createdAt).toLocaleDateString('en-IN', { dateStyle: 'long' })}</p>
        </div>
        
        <p>You can also download the Excel version from your dashboard.</p>
        
        <div style="text-align: center; margin: 30px 0;">
          <a href="${frontendUrl}/reports" 
             style="background: linear-gradient(135deg, #8B5CF6, #7C3AED); color: white; padding: 14px 28px; text-decoration: none; border-radius: 8px; display: inline-block; font-weight: 600;">
            View in Dashboard
          </a>
        </div>
        
        <hr style="border: none; border-top: 1px solid #E5E7EB; margin: 30px 0;">
        <p style="color: #6B7280; font-size: 12px; text-align: center;">
          Thank you for using CA Excel Report Generation Service!
        </p>
      </div>
    `,
    attachments: pdfBuffer ? [{
      filename: `${reportDetails.title || 'Report'}.pdf`,
      content: pdfBuffer,
      contentType: 'application/pdf'
    }] : []
  };
  
  await transporter.sendMail(mailOptions);
};

const sendInvoiceEmail = async (email, name, report, payment) => {
  // Skip sending email for test accounts
  if (isTestEmail(email)) {
    console.log(`Skipping invoice email for test account: ${email}`);
    return;
  }

  try {
    // Generate invoice PDF using the Python script
    const { spawn } = require('child_process');
    const path = require('path');
    const fs = require('fs').promises;
    const crypto = require('crypto');

    // Create temp directory if it doesn't exist
    const tempDir = path.join(__dirname, '../temp');
    try {
      await fs.access(tempDir);
    } catch {
      await fs.mkdir(tempDir, { recursive: true });
    }

    // Generate unique filename
    const invoiceId = crypto.randomBytes(8).toString('hex');
    const invoicePath = path.join(tempDir, `invoice-${invoiceId}.pdf`);

    // Prepare payment data for Python script
    const paymentData = {
      razorpay_order_id: payment.razorpay_order_id,
      razorpay_payment_id: payment.razorpay_payment_id,
      amount: payment.amount,
      currency: payment.currency,
      status: payment.status,
      paid_at: payment.paid_at
    };

    // Call Python script to generate invoice
    const pythonProcess = spawn('python', [
      path.join(__dirname, '../python-engine/pdf_generator.py'),
      'invoice',
      JSON.stringify(paymentData),
      invoicePath
    ]);

    return new Promise((resolve, reject) => {
      pythonProcess.on('close', async (code) => {
        if (code === 0) {
          try {
            // Read the generated PDF
            const pdfBuffer = await fs.readFile(invoicePath);

            // Send email with invoice attachment
            const mailOptions = {
              from: process.env.EMAIL_USER,
              to: email,
              subject: `Invoice for Your Report - ${report.title}`,
              html: `
                <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px;">
                  <h2 style="color: #7C3AED;">Payment Successful!</h2>
                  <p>Dear ${name},</p>
                  <p>Thank you for your payment. Your report has been processed successfully.</p>

                  <div style="background-color: #F3F4F6; padding: 20px; margin: 20px 0; border-radius: 8px;">
                    <h3 style="margin-top: 0; color: #374151;">Payment Details:</h3>
                    <ul style="list-style: none; padding: 0;">
                      <li style="margin: 8px 0;"><strong>Report:</strong> ${report.title}</li>
                      <li style="margin: 8px 0;"><strong>Amount:</strong> ₹${payment.amount.toLocaleString('en-IN')}</li>
                      <li style="margin: 8px 0;"><strong>Payment ID:</strong> ${payment.razorpay_payment_id}</li>
                      <li style="margin: 8px 0;"><strong>Date:</strong> ${new Date(payment.paid_at).toLocaleDateString('en-IN', { dateStyle: 'long' })}</li>
                    </ul>
                  </div>

                  <p>Your invoice is attached to this email. You can also download your report from your dashboard.</p>
                  <p>If you have any questions, please contact our support team.</p>

                  <div style="text-align: center; margin: 30px 0;">
                    <a href="${process.env.FRONTEND_URL || 'http://localhost:5173'}/reports"
                       style="background: linear-gradient(135deg, #8B5CF6, #7C3AED); color: white; padding: 14px 28px; text-decoration: none; border-radius: 8px; display: inline-block; font-weight: 600;">
                      View Reports
                    </a>
                  </div>

                  <hr style="border: none; border-top: 1px solid #E5E7EB; margin: 30px 0;">
                  <p style="color: #6B7280; font-size: 12px; text-align: center;">
                    Thank you for using CA Excel Report Generation Service!
                  </p>
                </div>
              `,
              attachments: [{
                filename: `invoice-${report.title}.pdf`,
                content: pdfBuffer,
                contentType: 'application/pdf'
              }]
            };

            await transporter.sendMail(mailOptions);
            console.log(`Invoice email sent successfully to ${email}`);

            // Clean up temp file
            fs.unlink(invoicePath).catch(console.error);

            resolve();
          } catch (error) {
            console.error('Error sending invoice email:', error);
            reject(error);
          }
        } else {
          console.error('Python script failed to generate invoice for email');
          reject(new Error('Failed to generate invoice PDF'));
        }
      });

      pythonProcess.on('error', (error) => {
        console.error('Error running Python script for invoice email:', error);
        reject(error);
      });
    });

  } catch (error) {
    console.error('Error in sendInvoiceEmail:', error);
    throw error;
  }
};

module.exports = {
  sendVerificationEmail,
  sendWelcomeEmail,
  sendOrderConfirmation,
  sendPasswordResetOTP,
  sendReportGeneratedNotification,
  sendWithdrawalStatusNotification,
  sendReportApprovalEmail,
  sendReportRejectionEmail,
  sendReportWithAttachment,
  sendInvoiceEmail
};