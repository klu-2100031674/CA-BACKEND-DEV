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
    const formData = req.body;
    const grokApiKey = req.body.grokApiKey || req.body.apiKey || process.env.GROK_API_KEY || process.env.XAI_API_KEY;

    logger.business('Generating full AI-enhanced report (Grok-only)', {
      userId: req.user ? req.user._id : null,
      templateId,
      operation: 'downloadFullReport',
      creditsRequired: 100
    });

    // Validate Grok API key first
    if (!grokApiKey) {
      return res.status(400).json({
        success: false,
        error: 'Grok API key required. Provide grokApiKey in request body or set GROK_API_KEY/XAI_API_KEY environment variable.'
      });
    }

    // Check wallet credits (100 credits for AI report)
    const wallet = await Wallet.findOne({ user_id: req.user._id });
    if (!wallet || wallet.report_credits < 100) {
      return res.status(400).json({
        success: false,
        error: 'Insufficient report credits. AI report generation requires 100 credits.',
        error_code: 'INSUFFICIENT_REPORT_CREDITS',
        required_credits: 100,
        available_credits: wallet ? wallet.report_credits : 0
      });
    }

    // Deduct 100 credits from wallet immediately after validation
    wallet.report_credits -= 100;
    await wallet.save();

    logger.business('Credits deducted for AI report generation', {
      userId: req.user._id,
      templateId,
      creditsDeducted: 100,
      remainingCredits: wallet.report_credits,
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
        'grok'
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
        grokApiKey
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

    logger.business('AI report generated and saved successfully', {
      userId: req.user._id,
      templateId,
      fileId: excelFile._id,
      creditsDeducted: 100,
      remainingCredits: wallet.report_credits,
      usedExistingExcel,
      operation: 'downloadFullReport'
    });

    // Return success response with file ID and PDF download data only
    res.json({
      success: true,
      message: usedExistingExcel 
        ? 'AI-enhanced full report generated successfully using existing Excel data'
        : 'AI-enhanced full report generated successfully',
      data: {
        fileId: excelFile._id,
        fullReportFileName: result.fullReportFileName,
        fullReportUrl: `/api/reports/download/${excelFile._id}/ai-report`
      }
    });

  } catch (error) {
    // Refund credits if they were deducted but report generation failed
    try {
      if (wallet && wallet.report_credits !== undefined) {
        wallet.report_credits += 100;
        await wallet.save();
        logger.warn('Credits refunded due to report generation failure', {
          userId: req.user ? req.user._id : null,
          templateId: req.params.templateId,
          creditsRefunded: 100,
          newBalance: wallet.report_credits,
          operation: 'downloadFullReport'
        });
      }
    } catch (refundError) {
      logger.error('Failed to refund credits after report generation error', {
        userId: req.user ? req.user._id : null,
        templateId: req.params.templateId,
        refundError: refundError.message,
        originalError: error.message,
        operation: 'downloadFullReport'
      });
    }

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

