const express = require('express');
const router = express.Router();
const multer = require('multer');
const Report = require('../models/Report');
const User = require('../models/User');
const Notification = require('../models/Notification');
const { verifyToken, requireRole } = require('../middleware/auth');
const mailService = require('../services/mailService');
const r2Service = require('../services/cloudflareR2Service');
const pdfGenerationService = require('../services/pdfGenerationService');
const path = require('path');
const fs = require('fs');

// Configure multer for Excel file uploads (25MB limit)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 25 * 1024 * 1024 // 25 MB
  },
  fileFilter: (req, file, cb) => {
    // Only allow Excel files
    const allowedMimes = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
      'application/vnd.ms-excel', // .xls
      'application/octet-stream' // Some browsers may send this
    ];
    const allowedExts = ['.xlsx', '.xls'];
    const ext = path.extname(file.originalname).toLowerCase();
    
    if (allowedMimes.includes(file.mimetype) || allowedExts.includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error('Only Excel files (.xlsx, .xls) are allowed'), false);
    }
  }
});

// ============================================================================
// ADMIN REPORT MANAGEMENT ROUTES
// ============================================================================

/**
 * Get admin reports dashboard statistics
 * GET /admin-reports/stats
 */
router.get('/stats', verifyToken, requireRole(['admin', 'super_admin']), async (req, res) => {
  try {
    const [
      pendingCount,
      underReviewCount,
      approvedCount,
      rejectedCount,
      totalCount,
      todayCount
    ] = await Promise.all([
      Report.countDocuments({ validation_status: 'pending_validation' }),
      Report.countDocuments({ validation_status: 'under_review' }),
      Report.countDocuments({ validation_status: 'approved' }),
      Report.countDocuments({ validation_status: 'rejected' }),
      Report.countDocuments({ validation_status: { $ne: 'draft' } }),
      Report.countDocuments({
        validation_status: 'pending_validation',
        createdAt: { $gte: new Date().setHours(0, 0, 0, 0) }
      })
    ]);

    res.json({
      success: true,
      data: {
        pending: pendingCount,
        under_review: underReviewCount,
        approved: approvedCount,
        rejected: rejectedCount,
        total: totalCount,
        today_pending: todayCount
      }
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

/**
 * Get all reports for admin with filters
 * GET /admin-reports
 */
router.get('/', verifyToken, requireRole(['admin', 'super_admin']), async (req, res) => {
  try {
    const {
      status,
      report_type,
      user_id,
      search,
      start_date,
      end_date,
      page = 1,
      limit = 20,
      sort = '-createdAt'
    } = req.query;

    // Build filter
    const filter = {};
    
    // Exclude drafts by default
    if (status) {
      filter.validation_status = status;
    } else {
      filter.validation_status = { $ne: 'draft' };
    }

    if (report_type) filter.report_type = report_type;
    if (user_id) filter.user_id = user_id;
    
    if (search) {
      filter.$or = [
        { title: { $regex: search, $options: 'i' } },
        { client_name: { $regex: search, $options: 'i' } },
        { templateId: { $regex: search, $options: 'i' } }
      ];
    }

    if (start_date || end_date) {
      filter.createdAt = {};
      if (start_date) filter.createdAt.$gte = new Date(start_date);
      if (end_date) filter.createdAt.$lte = new Date(end_date);
    }

    const skip = (parseInt(page) - 1) * parseInt(limit);

    const [reports, total] = await Promise.all([
      Report.find(filter)
        .populate('user_id', 'name email phone role')
        .populate('validated_by', 'name email')
        .select('_id title templateId validation_status createdAt updatedAt payment report_type user_id validated_by approval_email_sent excel_file_url pdf_file_url json_file_url requested_sheets analysis_options client_name client_details rejection_reason')
        .sort(sort)
        .skip(skip)
        .limit(parseInt(limit))
        .lean(),
      Report.countDocuments(filter)
    ]);

    // Process reports to add signed URLs for R2 storage
    const processedReports = await Promise.all(reports.map(async (report) => {
      const reportObj = { ...report };

      // Generate signed URLs for R2 files
      if (reportObj.excel_file_url && reportObj.excel_file_url.includes('r2.cloudflarestorage.com')) {
        const key = r2Service.extractKeyFromUrl(reportObj.excel_file_url);
        if (key) {
          reportObj.excel_file_url = await r2Service.generatePresignedUrl(key);
        }
      }
      
      if (reportObj.pdf_file_url && reportObj.pdf_file_url.includes('r2.cloudflarestorage.com')) {
        const key = r2Service.extractKeyFromUrl(reportObj.pdf_file_url);
        if (key) {
          reportObj.pdf_file_url = await r2Service.generatePresignedUrl(key);
        }
      }

      if (reportObj.json_file_url && reportObj.json_file_url.includes('r2.cloudflarestorage.com')) {
        const key = r2Service.extractKeyFromUrl(reportObj.json_file_url);
        if (key) {
          reportObj.json_file_url = await r2Service.generatePresignedUrl(key);
        }
      }

      return reportObj;
    }));

    res.json({
      success: true,
      data: {
        reports: processedReports,
        pagination: {
          current_page: parseInt(page),
          total_pages: Math.ceil(total / parseInt(limit)),
          total_count: total,
          per_page: parseInt(limit)
        }
      }
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

/**
 * Get pending validation reports
 * GET /admin-reports/pending
 */
router.get('/pending', verifyToken, requireRole(['admin', 'super_admin']), async (req, res) => {
  try {
    const { page = 1, limit = 20 } = req.query;
    const skip = (parseInt(page) - 1) * parseInt(limit);

    const [reports, total] = await Promise.all([
      Report.find({ validation_status: 'pending_validation' })
        .populate('user_id', 'name email phone role')
        .select('_id title templateId validation_status createdAt updatedAt payment report_type user_id approval_email_sent requested_sheets analysis_options client_name client_details')
        .sort('-createdAt')
        .skip(skip)
        .limit(parseInt(limit))
        .lean(),
      Report.countDocuments({ validation_status: 'pending_validation' })
    ]);

    res.json({
      success: true,
      data: {
        reports,
        pagination: {
          current_page: parseInt(page),
          total_pages: Math.ceil(total / parseInt(limit)),
          total_count: total
        }
      }
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

/**
 * Get payment analytics for reports
 * GET /admin-reports/payments
 */
router.get('/payments', verifyToken, requireRole(['admin', 'super_admin']), async (req, res) => {
  try {
    const { period = 'all', page = 1, limit = 50 } = req.query;
    const skip = (parseInt(page) - 1) * parseInt(limit);

    // Define date ranges
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const yesterday = new Date(today);
    yesterday.setDate(yesterday.getDate() - 1);
    const weekAgo = new Date(today);
    weekAgo.setDate(weekAgo.getDate() - 7);
    const monthAgo = new Date(today);
    monthAgo.setDate(monthAgo.getDate() - 30);

    let dateFilter = {};
    let periodLabel = 'All Time';

    switch (period) {
      case 'today':
        dateFilter = { 'payment.paid_at': { $gte: today } };
        periodLabel = 'Today';
        break;
      case 'yesterday':
        dateFilter = {
          'payment.paid_at': {
            $gte: yesterday,
            $lt: today
          }
        };
        periodLabel = 'Yesterday';
        break;
      case 'week':
        dateFilter = { 'payment.paid_at': { $gte: weekAgo } };
        periodLabel = 'Last 7 Days';
        break;
      case 'month':
        dateFilter = { 'payment.paid_at': { $gte: monthAgo } };
        periodLabel = 'Last 30 Days';
        break;
      default:
        // No date filter for 'all'
        break;
    }

    // Base filter for completed payments
    const baseFilter = {
      'payment.status': 'completed',
      ...dateFilter
    };

    // Get total count and stats
    const [totalCount, totalRevenue, todayRevenue, weekRevenue, monthRevenue] = await Promise.all([
      Report.countDocuments(baseFilter),
      Report.aggregate([
        { $match: { 'payment.status': 'completed' } },
        { $group: { _id: null, total: { $sum: '$payment.amount' } } }
      ]),
      Report.aggregate([
        { $match: { 'payment.status': 'completed', 'payment.paid_at': { $gte: today } } },
        { $group: { _id: null, total: { $sum: '$payment.amount' } } }
      ]),
      Report.aggregate([
        { $match: { 'payment.status': 'completed', 'payment.paid_at': { $gte: weekAgo } } },
        { $group: { _id: null, total: { $sum: '$payment.amount' } } }
      ]),
      Report.aggregate([
        { $match: { 'payment.status': 'completed', 'payment.paid_at': { $gte: monthAgo } } },
        { $group: { _id: null, total: { $sum: '$payment.amount' } } }
      ])
    ]);

    // Get paginated payments with user details
    const payments = await Report.find(baseFilter)
      .populate('user_id', 'name email')
      .select('_id title templateId payment user_id createdAt')
      .sort({ 'payment.paid_at': -1 })
      .skip(skip)
      .limit(parseInt(limit))
      .lean();

    // Group payments by date for better organization
    const groupedPayments = {};
    payments.forEach(payment => {
      const date = new Date(payment.payment.paid_at);
      const dateKey = date.toISOString().split('T')[0]; // YYYY-MM-DD format

      if (!groupedPayments[dateKey]) {
        groupedPayments[dateKey] = {
          date: dateKey,
          displayDate: date.toLocaleDateString('en-IN', {
            weekday: 'long',
            year: 'numeric',
            month: 'long',
            day: 'numeric'
          }),
          payments: [],
          totalAmount: 0,
          count: 0
        };
      }

      groupedPayments[dateKey].payments.push({
        id: payment._id,
        title: payment.title,
        templateId: payment.templateId,
        user: {
          name: payment.user_id?.name || 'Unknown',
          email: payment.user_id?.email || 'N/A'
        },
        amount: payment.payment.amount,
        currency: payment.payment.currency,
        paymentId: payment.payment.razorpay_payment_id,
        orderId: payment.payment.razorpay_order_id,
        paidAt: payment.payment.paid_at,
        createdAt: payment.createdAt
      });

      groupedPayments[dateKey].totalAmount += payment.payment.amount;
      groupedPayments[dateKey].count += 1;
    });

    // Convert to array and sort by date (newest first)
    const groupedPaymentsArray = Object.values(groupedPayments).sort((a, b) =>
      new Date(b.date) - new Date(a.date)
    );

    res.json({
      success: true,
      data: {
        period: periodLabel,
        summary: {
          total_payments: totalCount,
          total_revenue: totalRevenue[0]?.total || 0,
          today_revenue: todayRevenue[0]?.total || 0,
          week_revenue: weekRevenue[0]?.total || 0,
          month_revenue: monthRevenue[0]?.total || 0
        },
        payments: groupedPaymentsArray,
        pagination: {
          current_page: parseInt(page),
          total_pages: Math.ceil(totalCount / parseInt(limit)),
          total_count: totalCount,
          per_page: parseInt(limit)
        }
      },
      message: `Found ${totalCount} payments for ${periodLabel.toLowerCase()}`
    });

  } catch (error) {
    console.error('Error fetching payment analytics:', error);
    res.status(500).json({
      success: false,
      error: error.message,
      message: 'Failed to fetch payment analytics'
    });
  }
});

/**
 * Get single report details for admin review
 * GET /admin-reports/:id
 */
router.get('/:id', verifyToken, requireRole(['admin', 'super_admin']), async (req, res) => {
  try {
    const report = await Report.findById(req.params.id)
      .populate('user_id', 'name email phone role referral_code')
      .populate('validated_by', 'name email')
      .select('-excel_data') // Keep json_data for preview
      .lean();

    if (!report) {
      return res.status(404).json({ error: 'Report not found' });
    }

    const reportObj = { ...report };

    // Generate signed URLs for R2 files
    if (reportObj.excel_file_url && reportObj.excel_file_url.includes('r2.cloudflarestorage.com')) {
      const key = r2Service.extractKeyFromUrl(reportObj.excel_file_url);
      if (key) {
        reportObj.excel_file_url = await r2Service.generatePresignedUrl(key);
      }
    }
    
    if (reportObj.pdf_file_url && reportObj.pdf_file_url.includes('r2.cloudflarestorage.com')) {
      const key = r2Service.extractKeyFromUrl(reportObj.pdf_file_url);
      if (key) {
        reportObj.pdf_file_url = await r2Service.generatePresignedUrl(key);
      }
    }

    if (reportObj.json_file_url && reportObj.json_file_url.includes('r2.cloudflarestorage.com')) {
      const key = r2Service.extractKeyFromUrl(reportObj.json_file_url);
      if (key) {
        reportObj.json_file_url = await r2Service.generatePresignedUrl(key);
      }
    }

    res.json({
      success: true,
      data: reportObj
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

/**
 * Download report Excel file
 * GET /admin-reports/:id/excel
 */
router.get('/:id/excel', verifyToken, requireRole(['admin', 'super_admin']), async (req, res) => {
  try {
    const report = await Report.findById(req.params.id);
    
    if (!report) {
      return res.status(404).json({ error: 'Report not found' });
    }

    // If excel_file_url is R2 URL (cloud storage)
    if (report.excel_file_url && report.excel_file_url.includes('r2.cloudflarestorage.com')) {
      const r2Service = require('../services/cloudflareR2Service');
      const key = r2Service.extractKeyFromUrl(report.excel_file_url);
      
      if (key) {
        const fileBuffer = await r2Service.downloadFile(key);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${report.title || 'report'}.xlsx"`);
        return res.send(fileBuffer);
      }
    }

    // If excel_data is stored in DB (legacy/fallback)
    if (report.excel_data) {
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename="${report.title || 'report'}.xlsx"`);
      return res.send(report.excel_data);
    }

    // If excel_file_url is a local file path (legacy)
    if (report.excel_file_url) {
      const filePath = path.join(__dirname, '../../', report.excel_file_url);
      if (fs.existsSync(filePath)) {
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${report.title || 'report'}.xlsx"`);
        return res.sendFile(filePath);
      }
    }

    res.status(404).json({ 
      error: 'Excel file not found',
      message: 'This report was generated before cloud storage migration. Please ask the user to regenerate the report.',
      report_id: report._id
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

/**
 * Download/View report PDF file
 * GET /admin-reports/:id/pdf
 */
router.get('/:id/pdf', verifyToken, requireRole(['admin', 'super_admin']), async (req, res) => {
  try {
    const report = await Report.findById(req.params.id);
    
    if (!report) {
      return res.status(404).json({ error: 'Report not found' });
    }

    // If pdf_file_url is R2 URL (cloud storage)
    if (report.pdf_file_url && report.pdf_file_url.includes('r2.cloudflarestorage.com')) {
      const r2Service = require('../services/cloudflareR2Service');
      const key = r2Service.extractKeyFromUrl(report.pdf_file_url);
      
      if (key) {
        const fileBuffer = await r2Service.downloadFile(key);
        const inline = req.query.inline === 'true';
        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', `${inline ? 'inline' : 'attachment'}; filename="${report.title || 'report'}.pdf"`);
        return res.send(fileBuffer);
      }
    }

    // If pdf_file_url is a local file path (legacy)
    if (report.pdf_file_url) {
      const filePath = path.join(__dirname, '../../', report.pdf_file_url);
      if (fs.existsSync(filePath)) {
        const inline = req.query.inline === 'true';
        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', `${inline ? 'inline' : 'attachment'}; filename="${report.title || 'report'}.pdf"`);
        return res.sendFile(filePath);
      }
    }

    res.status(404).json({ 
      error: 'PDF file not found',
      message: 'This report was generated before cloud storage migration. Please ask the user to regenerate the report.',
      report_id: report._id
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

/**
 * Mark report as under review
 * PATCH /admin-reports/:id/review
 */
router.patch('/:id/review', verifyToken, requireRole(['admin', 'super_admin']), async (req, res) => {
  try {
    const report = await Report.findById(req.params.id);
    
    if (!report) {
      return res.status(404).json({ error: 'Report not found' });
    }

    if (report.validation_status !== 'pending_validation') {
      return res.status(400).json({ error: 'Report is not pending validation' });
    }

    report.validation_status = 'under_review';
    report.validated_by = req.user._id;
    await report.save();

    // Create notification for user
    await Notification.createNotification({
      user_id: report.user_id,
      type: 'report_under_review',
      title: 'Report Under Review',
      message: `Your report "${report.title}" is now being reviewed by our team.`,
      data: { report_id: report._id }
    });

    res.json({
      success: true,
      message: 'Report marked as under review',
      data: report
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

/**
 * Approve report
 * PATCH /admin-reports/:id/approve
 */
router.patch('/:id/approve', verifyToken, requireRole(['admin', 'super_admin']), async (req, res) => {
  try {
    const { validation_notes, send_email = true } = req.body;
    
    const report = await Report.findById(req.params.id).populate('user_id', 'name email');
    
    if (!report) {
      return res.status(404).json({ error: 'Report not found' });
    }

    if (!['pending_validation', 'under_review'].includes(report.validation_status)) {
      return res.status(400).json({ error: 'Report cannot be approved in current status' });
    }

    // Update report
    report.validation_status = 'approved';
    report.validated_by = req.user._id;
    report.validated_at = new Date();
    report.validation_notes = validation_notes;
    report.status = 'completed';

    // If admin has a signature, regenerate the PDF to include it
    if (req.user.signature_url) {
      try {
        console.log(`[Admin Approve] Admin has signature, regenerating PDF for report ${report._id}`);
        
        // 1. Download current Excel from R2
        const excelKey = r2Service.extractKeyFromUrl(report.excel_file_url);
        if (excelKey) {
          const excelBuffer = await r2Service.downloadFile(excelKey);
          
          // 2. Regenerate PDF
          const grokApiKey = process.env.GROK_API_KEY || process.env.XAI_API_KEY;
          const pdfResult = await pdfGenerationService.regeneratePdfFromExcel({
            excelBuffer: excelBuffer,
            templateId: report.templateId,
            selectedSheets: report.requested_sheets,
            grokApiKey: grokApiKey,
            jsonData: report.json_data,
            templateName: report.templateId,
            signatureUrl: req.user.signature_url
          });
          
          if (pdfResult.success) {
            // 3. Upload new PDF
            const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
            const pdfFileName = `${report.templateId}_approved_${timestamp}.pdf`;
            const userEmail = report.user_id?.email || report.user_id?._id?.toString() || 'unknown';
            
            const newPdfUrl = await r2Service.uploadPDF({
              fileBuffer: pdfResult.pdfBuffer,
              userEmail: userEmail,
              fileName: pdfFileName
            });
            
            // 4. Update report with new PDF URL
            report.pdf_file_url = newPdfUrl;
            console.log(`[Admin Approve] PDF regenerated and uploaded: ${newPdfUrl}`);
          }
        }
      } catch (regenError) {
        console.error(`[Admin Approve] Failed to regenerate PDF with signature:`, regenError);
        // Continue with approval even if regeneration fails
      }
    }

    await report.save();

    // Create notification
    await Notification.createNotification({
      user_id: report.user_id._id,
      type: 'report_approved',
      title: 'Report Approved! ✅',
      message: `Great news! Your report "${report.title}" has been approved. You can now download it.`,
      data: { report_id: report._id }
    });

    // Send email
    if (send_email && report.user_id?.email) {
      try {
        await mailService.sendReportApprovalEmail(
          report.user_id.email,
          report.user_id.name,
          {
            title: report.title,
            report_type: report.report_type || report.templateId,
            validation_notes: validation_notes
          }
        );
        report.approval_email_sent = true;
        await report.save();
      } catch (emailError) {
        console.error('Failed to send approval email:', emailError);
      }
    }

    res.json({
      success: true,
      message: 'Report approved successfully',
      data: report
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

/**
 * Reject report
 * PATCH /admin-reports/:id/reject
 */
router.patch('/:id/reject', verifyToken, requireRole(['admin', 'super_admin']), async (req, res) => {
  try {
    const { rejection_reason, validation_notes, send_email = true } = req.body;
    
    if (!rejection_reason) {
      return res.status(400).json({ error: 'Rejection reason is required' });
    }

    const report = await Report.findById(req.params.id).populate('user_id', 'name email');
    
    if (!report) {
      return res.status(404).json({ error: 'Report not found' });
    }

    if (!['pending_validation', 'under_review'].includes(report.validation_status)) {
      return res.status(400).json({ error: 'Report cannot be rejected in current status' });
    }

    // Update report
    report.validation_status = 'rejected';
    report.validated_by = req.user._id;
    report.validated_at = new Date();
    report.rejection_reason = rejection_reason;
    report.validation_notes = validation_notes;
    await report.save();

    // Create notification
    await Notification.createNotification({
      user_id: report.user_id._id,
      type: 'report_rejected',
      title: 'Report Needs Revision',
      message: `Your report "${report.title}" needs revision. Reason: ${rejection_reason}`,
      data: { report_id: report._id }
    });

    // Send email
    if (send_email && report.user_id?.email) {
      try {
        await mailService.sendReportRejectionEmail(
          report.user_id.email,
          report.user_id.name,
          {
            title: report.title,
            report_type: report.report_type || report.templateId,
            rejection_reason: rejection_reason
          }
        );
      } catch (emailError) {
        console.error('Failed to send rejection email:', emailError);
      }
    }

    res.json({
      success: true,
      message: 'Report rejected',
      data: report
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

/**
 * Upload revised Excel and regenerate PDF
 * POST /admin-reports/:id/upload-revised-excel
 * 
 * Only allowed when report is in 'under_review' status.
 * Admin must click "Start Review" first before uploading revised Excel.
 */
router.post('/:id/upload-revised-excel', 
  verifyToken, 
  requireRole(['admin', 'super_admin']), 
  upload.single('excel'),
  async (req, res) => {
    const reportId = req.params.id;
    const { revision_notes } = req.body;
    
    // Store original URLs for potential rollback
    let originalExcelUrl = null;
    let originalPdfUrl = null;
    let newExcelUrl = null;
    let newPdfUrl = null;
    
    try {
      // Validate file
      if (!req.file) {
        return res.status(400).json({ 
          success: false,
          error: 'Excel file is required' 
        });
      }

      const excelBuffer = req.file.buffer;
      
      // Find report
      const report = await Report.findById(reportId);
      
      if (!report) {
        return res.status(404).json({ 
          success: false,
          error: 'Report not found' 
        });
      }

      // Check status - must be under_review
      if (report.validation_status !== 'under_review') {
        return res.status(400).json({ 
          success: false,
          error: 'Report must be under review to upload revised Excel. Click "Start Review" first.' 
        });
      }

      // Store original URLs for revision history and potential rollback
      originalExcelUrl = report.excel_file_url;
      originalPdfUrl = report.pdf_file_url;

      console.log(`[Admin Upload] Starting revised Excel upload for report ${reportId}`);
      console.log(`[Admin Upload] Original Excel URL: ${originalExcelUrl}`);
      console.log(`[Admin Upload] Original PDF URL: ${originalPdfUrl}`);

      // Get user email for R2 path
      const user = await User.findById(report.user_id);
      const userEmail = user?.email || report.user_id.toString();

      // Step 1: Upload revised Excel to R2
      console.log(`[Admin Upload] Uploading revised Excel to R2...`);
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const excelFileName = `${report.templateId}_revised_${timestamp}.xlsx`;
      
      newExcelUrl = await r2Service.uploadExcel({
        fileBuffer: excelBuffer,
        userEmail: userEmail,
        fileName: excelFileName
      });
      
      console.log(`[Admin Upload] New Excel uploaded: ${newExcelUrl}`);

      // Step 2: Regenerate PDF from the new Excel (with AI content and proper sheet ordering)
      console.log(`[Admin Upload] Regenerating PDF from revised Excel...`);
      
      // Get Grok API key from environment (same as reportController)
      const grokApiKey = process.env.GROK_API_KEY || process.env.XAI_API_KEY;
      
      if (!grokApiKey) {
        console.warn(`[Admin Upload] No Grok API key found. PDF will be generated without AI content.`);
      }
      
      const pdfResult = await pdfGenerationService.regeneratePdfFromExcel({
        excelBuffer: excelBuffer,
        templateId: report.templateId,
        selectedSheets: report.requested_sheets, // Use stored sheets from original generation
        grokApiKey: grokApiKey,
        jsonData: report.json_data, // Pass stored JSON data for AI context
        htmlData: null, // HTML data is not typically stored
        templateName: report.templateId, // Use templateId as template name
        signatureUrl: req.user.signature_url // Pass admin signature URL
      });

      if (!pdfResult.success) {
        // PDF generation failed - attempt rollback
        console.error(`[Admin Upload] PDF generation failed: ${pdfResult.error}`);
        console.log(`[Admin Upload] New Excel is already uploaded. Keeping it but notifying admin.`);
        
        // Update report with new Excel but keep old PDF reference
        // Add revision entry noting the failure
        report.excel_file_url = newExcelUrl;
        report.revision_history = report.revision_history || [];
        report.revision_history.push({
          revised_at: new Date(),
          revised_by: req.user._id,
          revision_notes: `Excel uploaded but PDF regeneration failed: ${pdfResult.error}. Old PDF retained.`,
          old_excel_url: originalExcelUrl,
          old_pdf_url: originalPdfUrl
        });
        await report.save();

        return res.status(500).json({
          success: false,
          error: `PDF regeneration failed: ${pdfResult.error}`,
          partial_success: true,
          message: 'Excel was uploaded but PDF could not be regenerated. Please try regenerating PDF manually or upload a corrected Excel file.',
          data: {
            new_excel_url: newExcelUrl,
            current_pdf_url: originalPdfUrl
          }
        });
      }

      // Step 3: Upload regenerated PDF to R2
      console.log(`[Admin Upload] Uploading regenerated PDF to R2...`);
      const pdfFileName = pdfResult.pdfFileName || `${report.templateId}_revised_${timestamp}.pdf`;
      
      newPdfUrl = await r2Service.uploadPDF({
        fileBuffer: pdfResult.pdfBuffer,
        userEmail: userEmail,
        fileName: pdfFileName
      });
      
      console.log(`[Admin Upload] New PDF uploaded: ${newPdfUrl}`);

      // Step 4: Update report with new URLs and revision history
      report.excel_file_url = newExcelUrl;
      report.pdf_file_url = newPdfUrl;
      
      // Add to revision history
      report.revision_history = report.revision_history || [];
      report.revision_history.push({
        revised_at: new Date(),
        revised_by: req.user._id,
        revision_notes: revision_notes || 'Admin uploaded revised Excel',
        old_excel_url: originalExcelUrl,
        old_pdf_url: originalPdfUrl
      });

      await report.save();

      console.log(`[Admin Upload] Report ${reportId} updated successfully`);

      // Create notification for user
      await Notification.createNotification({
        user_id: report.user_id,
        type: 'report_revised',
        title: 'Report Updated',
        message: `Your report "${report.title}" has been revised by admin.`,
        data: { report_id: report._id }
      });

      res.json({
        success: true,
        message: 'Revised Excel uploaded and PDF regenerated successfully',
        data: {
          report_id: report._id,
          new_excel_url: newExcelUrl,
          new_pdf_url: newPdfUrl,
          sheets_processed: pdfResult.sheetsProcessed,
          revision_count: report.revision_history.length
        }
      });

    } catch (error) {
      console.error(`[Admin Upload] Error: ${error.message}`);
      console.error(error.stack);
      
      res.status(500).json({ 
        success: false,
        error: error.message 
      });
    }
  }
);

/**
 * Bulk approve reports
 * POST /admin-reports/bulk-approve
 */
router.post('/bulk-approve', verifyToken, requireRole(['admin', 'super_admin']), async (req, res) => {
  try {
    const { report_ids, validation_notes, send_email = true } = req.body;
    
    if (!report_ids || !Array.isArray(report_ids) || report_ids.length === 0) {
      return res.status(400).json({ error: 'Report IDs array is required' });
    }

    const results = { success: 0, failed: 0, errors: [] };

    for (const reportId of report_ids) {
      try {
        const report = await Report.findById(reportId).populate('user_id', 'name email');
        
        if (!report || !['pending_validation', 'under_review'].includes(report.validation_status)) {
          results.failed++;
          results.errors.push({ id: reportId, error: 'Invalid report or status' });
          continue;
        }

        report.validation_status = 'approved';
        report.validated_by = req.user._id;
        report.validated_at = new Date();
        report.validation_notes = validation_notes;
        report.status = 'completed';
        await report.save();

        // Create notification
        await Notification.createNotification({
          user_id: report.user_id._id,
          type: 'report_approved',
          title: 'Report Approved! ✅',
          message: `Your report "${report.title}" has been approved.`,
          data: { report_id: report._id }
        });

        // Send email
        if (send_email && report.user_id?.email) {
          try {
            await mailService.sendReportApprovalEmail(
              report.user_id.email,
              report.user_id.name,
              { title: report.title, report_type: report.report_type || report.templateId }
            );
          } catch (e) { /* ignore email errors in bulk */ }
        }

        results.success++;
      } catch (err) {
        results.failed++;
        results.errors.push({ id: reportId, error: err.message });
      }
    }

    res.json({
      success: true,
      message: `Bulk approval completed: ${results.success} approved, ${results.failed} failed`,
      data: results
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

/**
 * Get report types for filter dropdown
 * GET /admin-reports/meta/report-types
 */
router.get('/meta/report-types', verifyToken, requireRole(['admin', 'super_admin']), async (req, res) => {
  try {
    const reportTypes = await Report.distinct('report_type');
    const templateIds = await Report.distinct('templateId');
    
    res.json({
      success: true,
      data: {
        report_types: reportTypes.filter(Boolean),
        template_ids: templateIds.filter(Boolean)
      }
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;
