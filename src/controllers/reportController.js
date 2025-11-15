const Report = require('../models/Report');
const Wallet = require('../models/Wallet');
const excelCalculationService = require('../services/excelCalculationService');
const fs = require('fs').promises;
const path = require('path');
const { spawn } = require('child_process');

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

    console.log(`[ReportController] Apply form data for template: ${templateId}`);
    console.log(`[ReportController] User: ${req.user._id}`);

    // Apply form data and calculate using Excel service
    const result = await excelCalculationService.applyFormDataAndCalculate(
      templateId,
      formData
    );

    // Return success response with file path and calculated data
    res.json({
      success: true,
      message: 'Excel generated successfully',
      data: {
        templateId: templateId,
        fileUrl: result.relativePath,
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
    console.error('[ReportController] Error in applyFormData:', error);
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
    console.error('[ReportController] Error in applyFinalEdits:', error);
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
    console.error('[ReportController] Error in exportToPdf:', error);
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
    console.error('[ReportController] Error in exportToJson:', error);
    next(error);
  }
};

/**
 * @route   GET /api/reports/templates/:templateId/download/:fileName
 * @desc    Download calculated Excel file
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
    console.error('[ReportController] Error in downloadExcelFile:', error);
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

    console.log(`[ReportController] Creating report: ${title}`);

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

    console.log(`[ReportController] Report created: ${report._id}`);

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
    console.error('[ReportController] Error in createReport:', error);
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
    console.error('[ReportController] Error in getReports:', error);
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
    console.error('[ReportController] Error in getReportById:', error);
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
    console.error('[ReportController] Error in uploadReportJson:', error);
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
    console.error('[ReportController] Error in deleteReport:', error);
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
        console.error(`Python script exited with code ${code}`);
        console.error(`stderr: ${stderr}`);
        return reject(new Error(`Python script failed with code ${code}`));
      }
      resolve(stdout);
    });

    pythonProcess.on('error', (err) => {
      console.error('Failed to start Python process.', err);
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
    const aiProvider = req.body.aiProvider || 'grok'; // Default to Grok (preferred AI provider)
    const apiKey = req.body.apiKey || req.body.perplexityApiKey || req.body.grokApiKey;

    console.log('[ReportController] Generating full AI-enhanced report (Python-only)');
    console.log(`  Template ID: ${templateId}`);
    console.log(`  AI Provider: ${aiProvider} (Grok is default)`);
    console.log(`  User: ${req.user ? req.user._id : 'Not authenticated'}`);

    // Validate API key based on provider
    let finalApiKey = apiKey;
    if (!finalApiKey) {
      if (aiProvider === 'grok') {
        finalApiKey = process.env.GROK_API_KEY;
      } else {
        finalApiKey = process.env.PERPLEXITY_API_KEY;
      }
    }

    if (!finalApiKey) {
      return res.status(400).json({
        success: false,
        error: `${aiProvider.toUpperCase()} API key required. Provide apiKey in request body or set ${aiProvider.toUpperCase()}_API_KEY environment variable. (Grok is the default AI provider)`
      });
    }

    // Use Excel calculation service with full report generation enabled
    const result = await excelCalculationService.generateFullReport(
      templateId,
      formData,
      finalApiKey,
      aiProvider
    );

    // Check if full report was generated
    if (!result.fullReportData) {
      return res.status(500).json({
        success: false,
        error: 'Full report generation failed. Check server logs for details.'
      });
    }

    // Return success response with all data
    res.json({
      success: true,
      message: 'AI-enhanced full report generated successfully',
      data: {
        // Excel data
        excelFileName: result.fileName,
        excelBase64: result.excelData,
        
        // FinalWorkings PDF
        pdfFileName: result.pdfFileName,
        pdfBase64: result.pdfData,
        
        // Full AI-enhanced report (Excel sheets + AI content)
        fullReportFileName: result.fullReportFileName,
        fullReportBase64: result.fullReportData,
        
        // JSON data for frontend display
        jsonData: result.jsonData,
        htmlContent: result.htmlContent,
        
        // Download URLs
        excelUrl: result.relativePath,
        fullReportUrl: `/temp/${result.fullReportFileName}`
      }
    });

  } catch (error) {
    console.error('[ReportController] Error in downloadFullReport:', error);
    next(error);
  }
};

module.exports = exports;

