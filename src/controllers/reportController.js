const Report = require('../models/Report');
const Wallet = require('../models/Wallet');
const excelCalculationService = require('../services/excelCalculationService');
const excelFileService = require('../services/excelFileService');
const fs = require('fs').promises;
const path = require('path');
const { spawn } = require('child_process');
const logger = require('../utils/logger');

/**
 * Report Controller
 * Handles all report-related operations including Excel generation
 */

/**
 * @route   POST /api/reports/templates/:templateId/apply-form
 * @desc    Apply form data to Excel template and get calculated results
 * @access  Private
 */
exports.applyFormData = async (req, res, next) => {
  try {
    const { templateId } = req.params;
    const formData = req.body;

    logger.business(`Applying form data for template: ${templateId}`, {
      userId: req.user._id,
      templateId,
      operation: 'applyFormData'
    });

    // Apply form data and calculate using Excel service
    const result = await excelCalculationService.applyFormDataAndCalculate(
      templateId,
      formData
    );

    // Save Excel file to MongoDB
    const excelFile = await excelFileService.saveExcelFile({
      userId: req.user._id,
      templateId: templateId,
      fileName: result.fileName,
      fileBuffer: Buffer.from(result.excelData, 'base64'),
      fileSize: Buffer.from(result.excelData, 'base64').length,
      generatedBy: req.user._id,
      stage: 'stage1',
      jsonData: result.jsonData,
      allSheetsData: result.allSheetsData,
      formattedWCData: result.formattedWCData,
      htmlContent: result.htmlContent,
      htmlJsonData: result.htmlJsonData,
      pdfData: result.pdfData ? Buffer.from(result.pdfData, 'base64') : null,
      pdfFileName: result.pdfFileName,
      meta: result.meta
    });

    // Return success response with file ID and calculated data
    res.json({
      success: true,
      message: 'Excel generated successfully',
      data: {
        fileId: excelFile._id,
        templateId: templateId,
        fileName: result.fileName,
        excelBase64: result.excelData, // Include the base64 Excel data
        jsonData: result.jsonData, // Include JSON data for Luckysheet
        allSheetsData: result.allSheetsData,
        formattedWCData: result.formattedWCData,
        htmlContent: result.htmlContent, // Include HTML content for frontend display
        htmlJsonData: result.htmlJsonData, // Include JSON data extracted from HTML
        pdfBase64: result.pdfData, // Include PDF data as fallback
        pdfFileName: result.pdfFileName, // Include PDF filename
        meta: result.meta
      }
    });

  } catch (error) {
    logger.error('Error in applyFormData', {
      userId: req.user._id,
      templateId: req.params.templateId,
      error: error.message,
      stack: error.stack,
      operation: 'applyFormData'
    });
    next(error);
  }
};

/**
 * @route   POST /api/reports/templates/:templateId/apply-final
 * @desc    Apply FinalWorkings edits (or any sheet updates) and return recalculated sheets
 * @access  Private
 */
exports.applyFinalEdits = async (req, res, next) => {
  try {
    const { templateId } = req.params;
    const { updates, recalculate } = req.body || {};

    if (!Array.isArray(updates) || updates.length === 0) {
      return res.status(400).json({ success: false, error: 'No updates provided' });
    }

    const result = await excelCalculationService.applyUpdatesAndCalculate(templateId, { updates, recalculate });

    res.json({
      success: true,
      message: 'Final edits applied successfully',
      data: {
        templateId: templateId,
        fileUrl: result.relativePath,
        fileName: result.fileName,
        excelBase64: result.excelData,
        jsonData: result.jsonData,
        allSheetsData: result.allSheetsData,
        formattedWCData: result.formattedWCData,
        htmlContent: result.htmlContent, // Include HTML content for frontend display
        htmlJsonData: result.htmlJsonData, // Include JSON data extracted from HTML
        pdfBase64: result.pdfData, // Include PDF data as fallback
        pdfFileName: result.pdfFileName, // Include PDF filename
        meta: result.meta
      }
    });
  } catch (error) {
    logger.error('Error in applyFinalEdits', {
      userId: req.user._id,
      templateId: req.params.templateId,
      error: error.message,
      stack: error.stack,
      operation: 'applyFinalEdits'
    });
    next(error);
  }
};

/**
 * @route   POST /api/reports/export/pdf
 * @desc    Export report data to PDF
 * @access  Private
 */
exports.exportToPdf = async (req, res, next) => {
  try {
    const { jsonData } = req.body;
    if (!jsonData) {
      return res.status(400).json({ success: false, error: 'No JSON data provided' });
    }

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const pdfFileName = `report-${timestamp}.pdf`;
    const pdfPath = path.join(__dirname, '../../temp', pdfFileName);

    const pythonEnginePath = path.join(__dirname, '../python-engine');
    const pythonExecutable = path.join(pythonEnginePath, '.venv/Scripts/python');
    const scriptPath = path.join(pythonEnginePath, 'pdf_generator.py');

    await runPythonScript(scriptPath, [JSON.stringify(jsonData), pdfPath], pythonExecutable);

    res.json({
      success: true,
      message: 'PDF generated successfully',
      data: {
        fileName: pdfFileName,
        url: `/api/reports/download/${pdfFileName}`
      }
    });

  } catch (error) {
    logger.error('Error in exportToPdf', {
      userId: req.user._id,
      error: error.message,
      stack: error.stack,
      operation: 'exportToPdf'
    });
    next(error);
  }
};

/**
 * @route   POST /api/reports/export/json
 * @desc    Export report data to a JSON file
 * @access  Private
 */
exports.exportToJson = async (req, res, next) => {
  try {
    const { jsonData } = req.body;
    if (!jsonData) {
      return res.status(400).json({ success: false, error: 'No JSON data provided' });
    }

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const jsonFileName = `report-${timestamp}.json`;
    const jsonPath = path.join(__dirname, '../../temp', jsonFileName);

    await fs.writeFile(jsonPath, JSON.stringify(jsonData, null, 2));

    res.json({
      success: true,
      message: 'JSON file generated successfully',
      data: {
        fileName: jsonFileName,
        url: `/api/reports/download/${jsonFileName}`
      }
    });

  } catch (error) {
    logger.error('Error in exportToJson', {
      userId: req.user._id,
      error: error.message,
      stack: error.stack,
      operation: 'exportToJson'
    });
    next(error);
  }
};

/**
 * @route   GET /api/reports/download/:fileId/:type
 * @desc    Download Excel file or AI report from MongoDB
 * @access  Private
 */
exports.downloadFile = async (req, res, next) => {
  try {
    const { fileId, type } = req.params; // type: 'excel' or 'ai-report'

    logger.business(`Downloading file: ${fileId}, type: ${type}`, {
      fileId,
      type,
      operation: 'downloadFile'
    });

    // Get file buffer from MongoDB
    const fileData = await excelFileService.getFileBuffer(fileId, type);

    // Increment download count
    await excelFileService.incrementDownloadCount(fileId);

    // Set headers for file download
    res.setHeader('Content-Type', fileData.mimeType);
    res.setHeader('Content-Disposition', `attachment; filename="${fileData.fileName}"`);
    res.setHeader('Content-Length', fileData.fileSize);

    // Send file buffer
    res.send(fileData.buffer);

  } catch (error) {
    logger.error('Error in downloadFile', {
      fileId: req.params.fileId,
      type: req.params.type,
      error: error.message,
      stack: error.stack,
      operation: 'downloadFile'
    });
    if (error.message.includes('not found')) {
      return res.status(404).json({
        success: false,
        error: 'File not found'
      });
    }
    next(error);
  }
};

/**
 * @route   GET /api/reports/templates/:templateId/download/:fileName (LEGACY)
 * @desc    Download calculated Excel file (legacy route for backwards compatibility)
 * @access  Private
 */
exports.downloadExcelFile = async (req, res, next) => {
  try {
    const { fileName } = req.params;
    const filePath = path.join(__dirname, '../../temp', fileName);

    // Check if file exists
    try {
      await fs.access(filePath);
    } catch (error) {
      return res.status(404).json({
        success: false,
        error: 'File not found or has expired'
      });
    }

    // Determine content type based on file extension
    const ext = path.extname(fileName).toLowerCase();
    let contentType = 'application/octet-stream';
    if (ext === '.xlsx') {
      contentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    } else if (ext === '.json') {
      contentType = 'application/json';
    } else if (ext === '.pdf') {
      contentType = 'application/pdf';
    }

    // Set headers for file download (force attachment)
    res.setHeader('Content-Type', contentType);
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);

    // Stream file to response
    const fileStream = require('fs').createReadStream(filePath);
    fileStream.pipe(res);

  } catch (error) {
    logger.error('Error in downloadExcelFile', {
      userId: req.user._id,
      templateId: req.params.templateId,
      fileName: req.params.fileName,
      error: error.message,
      stack: error.stack,
      operation: 'downloadExcelFile'
    });
    next(error);
  }
};

/**
 * @route   POST /api/reports
 * @desc    Create a new report record
 * @access  Private
 */
exports.createReport = async (req, res, next) => {
  try {
    const {
      title,
      templateId,
      excel_file_url,
      pdf_file_url,
      form_data,
      hidden_sheets,
      locked_sheets
    } = req.body;

    logger.business(`Creating report: ${title}`, {
      userId: req.user._id,
      title,
      templateId,
      operation: 'createReport'
    });

    // Check wallet credits
    const wallet = await Wallet.findOne({ user_id: req.user._id });
    if (!wallet || wallet.report_credits < 1) {
      return res.status(400).json({
        success: false,
        error: 'Insufficient report credits',
        error_code: 'INSUFFICIENT_REPORT_CREDITS'
      });
    }

    // Create report
    const report = new Report({
      user_id: req.user._id,
      title,
      templateId,
      excel_file_url,
      pdf_file_url,
      form_data,
      hidden_sheets: hidden_sheets || [],
      locked_sheets: locked_sheets || [],
      status: 'completed'
    });

    await report.save();

    // Deduct credit
    wallet.report_credits -= 1;
    await wallet.save();

    logger.business(`Report created successfully: ${report._id}`, {
      userId: req.user._id,
      reportId: report._id,
      title: report.title,
      templateId: report.templateId,
      creditsDeducted: 1,
      remainingCredits: wallet.report_credits,
      operation: 'createReport'
    });

    res.status(201).json({
      success: true,
      message: 'Report created successfully',
      data: {
        _id: report._id,
        title: report.title,
        templateId: report.templateId,
        status: report.status,
        hidden_sheets: report.hidden_sheets,
        locked_sheets: report.locked_sheets,
        createdAt: report.createdAt
      }
    });

  } catch (error) {
    logger.error('Error in createReport', {
      userId: req.user._id,
      title: req.body.title,
      templateId: req.body.templateId,
      error: error.message,
      stack: error.stack,
      operation: 'createReport'
    });
    next(error);
  }
};

/**
 * @route   GET /api/reports
 * @desc    Get all reports for current user
 * @access  Private
 */
exports.getReports = async (req, res, next) => {
  try {
    const query = req.user.role === 'admin'
      ? {}
      : { user_id: req.user._id };

    const reports = await Report.find(query)
      .populate('user_id', 'name email')
      .sort({ createdAt: -1 });

    res.json({
      success: true,
      count: reports.length,
      data: reports
    });

  } catch (error) {
    logger.error('Error in getReports', {
      userId: req.user._id,
      error: error.message,
      stack: error.stack,
      operation: 'getReports'
    });
    next(error);
  }
};

/**
 * @route   GET /api/reports/:reportId
 * @desc    Get single report by ID
 * @access  Private
 */
exports.getReportById = async (req, res, next) => {
  try {
    const { reportId } = req.params;

    const report = await Report.findById(reportId)
      .populate('user_id', 'name email');

    if (!report) {
      return res.status(404).json({
        success: false,
        error: 'Report not found'
      });
    }

    // Check ownership (unless admin)
    if (req.user.role !== 'admin' && report.user_id._id.toString() !== req.user._id.toString()) {
      return res.status(403).json({
        success: false,
        error: 'Not authorized to access this report'
      });
    }

    res.json({
      success: true,
      data: report
    });

  } catch (error) {
    logger.error('Error in getReportById', {
      userId: req.user._id,
      reportId: req.params.reportId,
      error: error.message,
      stack: error.stack,
      operation: 'getReportById'
    });
    next(error);
  }
};

/**
 * @route   POST /api/reports/:reportId/upload-json
 * @desc    Upload final JSON data for a report
 * @access  Public (for backwards compatibility)
 */
exports.uploadReportJson = async (req, res, next) => {
  try {
    const { reportId } = req.params;
    const { finalData } = req.body;

    const report = await Report.findById(reportId);

    if (!report) {
      return res.status(404).json({
        success: false,
        error: 'Report not found'
      });
    }

    if (!finalData) {
      return res.status(400).json({
        success: false,
        error: 'No final data provided'
      });
    }

    // Save JSON file
    const jsonFilePath = path.join(__dirname, '../../uploads', `${reportId}.json`);
    await fs.writeFile(jsonFilePath, JSON.stringify(finalData, null, 2));

    // Update report
    report.json_file_url = `/uploads/${reportId}.json`;
    await report.save();

    res.json({
      success: true,
      message: 'JSON uploaded successfully',
      data: {
        json_file_url: report.json_file_url
      }
    });

  } catch (error) {
    logger.error('Error in uploadReportJson', {
      reportId: req.params.reportId,
      error: error.message,
      stack: error.stack,
      operation: 'uploadReportJson'
    });
    next(error);
  }
};

/**
 * @route   DELETE /api/reports/:reportId
 * @desc    Delete a report
 * @access  Private
 */
exports.deleteReport = async (req, res, next) => {
  try {
    const { reportId } = req.params;

    const report = await Report.findById(reportId);

    if (!report) {
      return res.status(404).json({
        success: false,
        error: 'Report not found'
      });
    }

    // Check ownership (unless admin)
    if (req.user.role !== 'admin' && report.user_id.toString() !== req.user._id.toString()) {
      return res.status(403).json({
        success: false,
        error: 'Not authorized to delete this report'
      });
    }

    await report.deleteOne();

    res.json({
      success: true,
      message: 'Report deleted successfully'
    });

  } catch (error) {
    logger.error('Error in deleteReport', {
      userId: req.user._id,
      reportId: req.params.reportId,
      error: error.message,
      stack: error.stack,
      operation: 'deleteReport'
    });
    next(error);
  }
};

function runPythonScript(scriptPath, args, pythonExecutable) {
  return new Promise((resolve, reject) => {
    const pythonProcess = spawn(pythonExecutable, [scriptPath, ...args]);

    let stdout = '';
    let stderr = '';

    pythonProcess.stdout.on('data', (data) => {
      stdout += data.toString();
    });

    pythonProcess.stderr.on('data', (data) => {
      stderr += data.toString();
    });

    pythonProcess.on('close', (code) => {
      if (code !== 0) {
        logger.error('Python script execution failed', {
          scriptPath,
          exitCode: code,
          stderr,
          args,
          operation: 'runPythonScript'
        });
        return reject(new Error(`Python script failed with code ${code}`));
      }
      resolve(stdout);
    });

    pythonProcess.on('error', (err) => {
      logger.error('Failed to start Python process', {
        scriptPath,
        error: err.message,
        stack: err.stack,
        args,
        operation: 'runPythonScript'
      });
      reject(err);
    });
  });
}

/**
 * @route   POST /api/reports/templates/:templateId/download-full-report
 * @desc    Generate and download full AI-enhanced report (Excel PDFs + AI content) - Pure Python
 * @access  Private
 */
exports.downloadFullReport = async (req, res, next) => {
  try {
    const { templateId } = req.params;
    const requestPayload = req.body || {};
    const { selectedSheets, ...formData } = requestPayload;
    const normalizedSelectedSheets = Array.isArray(selectedSheets)
      ? selectedSheets
          .filter((sheet) => typeof sheet === 'string' && sheet.trim().length)
          .map((sheet) => sheet.trim())
      : null;
    const grokApiKey = requestPayload.grokApiKey || requestPayload.apiKey || process.env.GROK_API_KEY || process.env.XAI_API_KEY;

    logger.business('Generating full AI-enhanced report (Grok-only)', {
      userId: req.user ? req.user._id : null,
      templateId,
      operation: 'downloadFullReport',
      selectedSheets: normalizedSelectedSheets
    });

    // Validate Grok API key first
    if (!grokApiKey) {
      return res.status(400).json({
        success: false,
        error: 'Grok API key required. Provide grokApiKey in request body or set GROK_API_KEY/XAI_API_KEY environment variable.'
      });
    }

    // Check if payment has been made for this report
    // Look for a report with pending_validation status or paid payment status
    const Report = require('../models/Report');
    const report = await Report.findOne({
      user_id: req.user._id,
      templateId: templateId,
      'payment.status': 'completed',
      validation_status: { $in: ['pending_validation', 'draft'] }
    }).sort({ createdAt: -1 });

    if (!report) {
      return res.status(402).json({
        success: false,
        error: 'Payment required. Please complete payment before generating the final report.',
        error_code: 'PAYMENT_REQUIRED'
      });
    }

    logger.business('Payment verified for AI report generation', {
      userId: req.user._id,
      templateId,
      reportId: report._id,
      amountPaid: report.payment.amount,
      operation: 'downloadFullReport'
    });

    // Check for existing ReportStaging data (latest Excel with user edits)
    const ReportStaging = require('../models/ReportStaging');
    const existingStaging = await ReportStaging.findOne({
      user_id: req.user._id,
      template_id: templateId,
      status: 'active'
    }).sort({ createdAt: -1 });

    let result;
    let usedExistingExcel = false;

    if (existingStaging) {
      // Use existing Excel data from ReportStaging
      logger.business('Using existing Excel data from ReportStaging for AI report generation', {
        userId: req.user._id,
        templateId,
        stagingId: existingStaging._id,
        operation: 'downloadFullReport'
      });

      // Save Excel data to temporary file for processing
      const tempDir = path.join(__dirname, '../../temp');
      await fs.mkdir(tempDir, { recursive: true });
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const tempExcelPath = path.join(tempDir, `${templateId}_existing_${timestamp}.xlsx`);
      await fs.writeFile(tempExcelPath, existingStaging.excel_data);

      // Generate full report using existing Excel file
      result = await excelCalculationService.generateFullReportFromFile(
        tempExcelPath,
        templateId,
        formData, // Still pass formData for any additional context
        grokApiKey,
        'grok',
        { selectedSheets: normalizedSelectedSheets }
      );

      usedExistingExcel = true;

      // Cleanup temp file
      try {
        await fs.unlink(tempExcelPath);
      } catch (cleanupError) {
        logger.warn('Failed to cleanup temp Excel file', {
          tempExcelPath,
          error: cleanupError.message,
          operation: 'downloadFullReport'
        });
      }
    } else {
      // No existing staging data, generate fresh Excel
      logger.business('No existing Excel data found, generating fresh Excel for AI report', {
        userId: req.user._id,
        templateId,
        operation: 'downloadFullReport'
      });

      result = await excelCalculationService.generateFullReport(
        templateId,
        formData,
        grokApiKey,
        { selectedSheets: normalizedSelectedSheets }
      );
    }

    // Check if full report was generated
    if (!result.fullReportData) {
      return res.status(500).json({
        success: false,
        error: 'Full report generation failed. Check server logs for details.'
      });
    }

    // Save Excel file to MongoDB with AI report
    const excelFile = await excelFileService.saveExcelFile({
      userId: req.user._id,
      templateId: templateId,
      fileName: result.fileName,
      fileBuffer: Buffer.from(result.excelData, 'base64'),
      fileSize: Buffer.from(result.excelData, 'base64').length,
      generatedBy: req.user._id,
      stage: 'stage1',
      jsonData: result.jsonData,
      allSheetsData: result.allSheetsData,
      formattedWCData: result.formattedWCData,
      // htmlContent and htmlJsonData skipped (not generated)
      pdfData: result.pdfData ? Buffer.from(result.pdfData, 'base64') : null,
      pdfFileName: result.pdfFileName,
      meta: result.meta
    });

    // Update with AI report data
    await excelFileService.updateWithAIReport(excelFile._id, {
      reportBuffer: Buffer.from(result.fullReportData, 'base64'),
      reportFileName: result.fullReportFileName
    });

    // Upload Excel and PDF to Cloudflare R2
    const r2Service = require('../services/cloudflareR2Service');
    const userEmail = req.user.email || req.user._id.toString();
    
    try {
      // Upload Excel file to R2
      const excelUrl = await r2Service.uploadExcel({
        fileBuffer: Buffer.from(result.excelData, 'base64'),
        userEmail: userEmail,
        fileName: result.fileName,
      });

      // Upload PDF file to R2
      const pdfUrl = await r2Service.uploadPDF({
        fileBuffer: Buffer.from(result.fullReportData, 'base64'),
        userEmail: userEmail,
        fileName: result.fullReportFileName,
      });

      // Save R2 URLs to both Report and ExcelFile models
      if (!report) {
        throw new Error('Report not found - cannot save R2 URLs');
      }

      // Update Report model with R2 URLs
      report.excel_file_id = excelFile._id;
      report.excel_file_url = excelUrl; // R2 URL
      report.pdf_file_url = pdfUrl; // R2 URL
      
      // Update ExcelFile model with R2 URLs
      excelFile.excel_r2_url = excelUrl;
      excelFile.pdf_r2_url = pdfUrl;
      
      logger.info('Saving R2 URLs to report and excelFile', {
        userId: req.user._id,
        reportId: report._id,
        excelFileId: excelFile._id,
        excelUrl,
        pdfUrl,
      });

      await Promise.all([report.save(), excelFile.save()]);

      logger.info('✅ Files uploaded to Cloudflare R2 and URLs saved to database', {
        userId: req.user._id,
        reportId: report._id,
        excelFileId: excelFile._id,
        excelUrl,
        pdfUrl,
      });
    } catch (r2Error) {
      logger.error('R2 upload failed, falling back to local storage', {
        error: r2Error.message,
        reportId: report._id,
      });
      // Fallback: save to database as before
      if (report) {
        report.excel_file_id = excelFile._id;
        report.excel_data = Buffer.from(result.excelData, 'base64');
        await report.save();
      }
    }

    logger.business('AI report generated and saved successfully', {
      userId: req.user._id,
      templateId,
      fileId: excelFile._id,
      reportId: report._id,
      amountPaid: report.payment.amount,
      usedExistingExcel,
      operation: 'downloadFullReport'
    });

    // Return success response - report submitted for validation
    // Do NOT send file URLs or download links to frontend
    res.json({
      success: true,
      message: 'Report generated and submitted for validation',
      data: {
        report_id: report._id,
        validation_status: report.validation_status,
        message: 'Your report has been submitted for validation. You will receive an email notification once it is approved.'
      }
    });

  } catch (error) {
    // Note: No credit refund needed with pay-per-report model
    // Payment is verified before report generation starts
    
    logger.error('Error in downloadFullReport', {
      userId: req.user ? req.user._id : null,
      templateId: req.params.templateId,
      error: error.message,
      stack: error.stack,
      operation: 'downloadFullReport'
    });
    next(error);
  }
};

module.exports = exports;

