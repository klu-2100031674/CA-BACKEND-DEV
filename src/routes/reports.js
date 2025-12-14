const express = require('express');
const fs = require('fs').promises;
const path = require('path');
const zlib = require('zlib');
const Report = require('../models/Report');
const ReportStaging = require('../models/ReportStaging');
const Wallet = require('../models/Wallet');
const { verifyToken } = require('../middleware/auth');
const { alterTemplateJson } = require('../services/jsonAlterService');
const excelCalculationService = require('../services/excelCalculationService');
const reportController = require('../controllers/reportController');
const templateService = require('../services/templateService');
const logger = require('../utils/logger');
const router = express.Router();

logger.info('Reports route module loaded');

// Use template service for loading templates
async function loadTemplates() {
  return await templateService.loadTemplates();
}

router.get('/templates', async (req, res) => {
  try {
    const templates = await loadTemplates();
    const { search, author, page = 1, limit = 20, ...propertyFilters } = req.query;
    
    let filtered = [...templates];
    
    if (search) {
      const searchTerm = search.toLowerCase();
      filtered = filtered.filter(template => 
        template.name.toLowerCase().includes(searchTerm) ||
        template.description.toLowerCase().includes(searchTerm) ||
        template.id.toLowerCase().includes(searchTerm)
      );
    }
    
    if (author) {
      filtered = filtered.filter(template => template.author === author);
    }
    
    // Apply dynamic property filters
    Object.keys(propertyFilters).forEach(filterKey => {
      const filterValue = propertyFilters[filterKey];
      if (filterValue) {
        filtered = filtered.filter(template => {
          const templateValue = template.properties?.[filterKey];
          // Handle both string and number comparisons
          return templateValue !== undefined && 
                 (templateValue.toString() === filterValue || templateValue === parseInt(filterValue) || templateValue === filterValue);
        });
      }
    });
    
    const startIndex = (parseInt(page) - 1) * parseInt(limit);
    const endIndex = startIndex + parseInt(limit);
    const paginatedTemplates = filtered.slice(startIndex, endIndex);
    
    // Extract all unique properties and their values dynamically
    const allProperties = {};
    const authors = new Set();
    
    templates.forEach(template => {
      if (template.author) authors.add(template.author);
      
      if (template.properties) {
        Object.entries(template.properties).forEach(([key, value]) => {
          if (!allProperties[key]) {
            allProperties[key] = new Set();
          }
          allProperties[key].add(value);
        });
      }
    });
    
    // Convert sets to sorted arrays
    const dynamicFilters = {};
    Object.keys(allProperties).forEach(key => {
      const values = Array.from(allProperties[key]);
      // Sort numbers numerically, strings alphabetically
      dynamicFilters[key] = values.sort((a, b) => {
        if (typeof a === 'number' && typeof b === 'number') {
          return a - b;
        }
        return a.toString().localeCompare(b.toString());
      });
    });
    
    res.json({
      success: true,
      data: {
        templates: paginatedTemplates,
        total: filtered.length,
        page: parseInt(page),
        limit: parseInt(limit),
        totalPages: Math.ceil(filtered.length / parseInt(limit)),
        filters: {
          authors: Array.from(authors).sort(),
          properties: dynamicFilters
        }
      },
      message: `Found ${filtered.length} templates`
    });
  } catch (error) {
    res.status(500).json({ 
      success: false,
      error: error.message,
      message: 'Failed to load templates'
    });
  }
});

// Get template form HTML
router.get('/templates/:templateId/form', async (req, res) => {
  try {
    const { templateId } = req.params;
    
    try {
      const formPath = path.join(__dirname, `../../templates/forms/${templateId}.html`);
      logger.debug('Loading template form from file', { templateId, formPath });
      const formHtml = await fs.readFile(formPath, 'utf8');
      
      res.json({
        success: true,
        data: {
          html: formHtml,
          templateId: templateId
        },
        message: 'Template form retrieved successfully'
      });
    } catch (fileError) {
      logger.error('Template form file not found', {
        templateId,
        formPath: path.join(__dirname, `../../templates/forms/${templateId}.html`),
        error: fileError.message,
        operation: 'getTemplateForm'
      });
      res.status(404).json({ 
        success: false,
        error: 'Form HTML file not found',
        message: `Form for template ${templateId} not found`
      });
    }
  } catch (error) {
    res.status(500).json({ 
      success: false,
      error: error.message,
      message: 'Failed to retrieve template form'
    });
  }
});

router.get('/templates/:templateId', async (req, res) => {
  try {
    const { templateId } = req.params;
    const templates = await loadTemplates();
    
    const template = templates.find(t => t.id === templateId);
    if (!template) {
      return res.status(404).json({ 
        success: false,
        error: 'Template not found',
        message: `Template with ID ${templateId} not found`
      });
    }
    
    // Only return metadata, no JSON data
    res.json({
      success: true,
      data: template,
      message: 'Template retrieved successfully'
    });
  } catch (error) {
    res.status(500).json({ 
      success: false,
      error: error.message,
      message: 'Failed to retrieve template'
    });
  }
});

// New endpoint to get template with form data applied
// POST /api/reports/templates/:templateId/apply-form - Apply form data to Excel template (NEW PROPER APPROACH)
router.post('/templates/:templateId/apply-form', verifyToken, async (req, res) => {
  try {
    const { templateId } = req.params;
    const formData = req.body;
    
    logger.business('Processing template with form data', {
      userId: req.user._id,
      templateId,
      formDataKeys: Object.keys(formData).length,
      operation: 'applyFormToTemplate'
    });
    
    // Validate template exists
    const templates = await loadTemplates();
    const template = templates.find(t => t.id === templateId);
    if (!template) {
      return res.status(404).json({ error: 'Template not found' });
    }

    try {
      // 🎯 YOUR DESIRED APPROACH:
      // 1. Take form data from frontend ✅
      // 2. Send data to Excel and insert in Assumptions sheet ✅ 
      // 3. Recalculate the Excel ✅
      // 4. Get all data from multiple sheets in proper format ✅
      
      logger.debug('Processing form data for Excel calculation', {
        templateId,
        userId: req.user._id,
        payloadSize: JSON.stringify(formData).length,
        operation: 'applyFormDataAndCalculate'
      });
      
      // ⚠️ IMPORTANT: Pass the ORIGINAL formData to the service
      // The service's extractFormData() will handle unwrapping for cell data
      // The service's extractFixedAssetsSchedule() needs the full payload with additionalData
      const result = await excelCalculationService.applyFormDataAndCalculate(templateId, formData);

      logger.business('Excel calculation completed successfully', {
        templateId,
        userId: req.user._id,
        operation: 'applyFormDataAndCalculate'
      });

      // Save Excel to staging for potential full report generation
      try {
        // Clean up old staging records for this user/template (keep only the latest)
        await ReportStaging.deleteMany({
          user_id: req.user._id,
          template_id: templateId,
          status: 'active'
        });

        const stagingRecord = new ReportStaging({
          user_id: req.user._id,
          template_id: templateId,
          excel_data: Buffer.from(result.excelData, 'base64'),
          file_name: result.fileName
        });
        await stagingRecord.save();
        logger.business('Excel data saved to staging', {
          userId: req.user._id,
          templateId,
          stagingId: stagingRecord._id,
          fileName: result.fileName,
          operation: 'saveToStaging'
        });
      } catch (stagingError) {
        logger.error('Failed to save Excel to staging', {
          userId: req.user._id,
          templateId,
          error: stagingError.message,
          operation: 'saveToStaging'
        });
        // Continue anyway
      }

      // Return Excel filename and base64 data for frontend to display
      res.json({
        success: true,
        message: 'Excel, PDF, and HTML generated successfully',
        data: {
          fileName: result.fileName,
          excelBase64: result.excelData,
          pdfBase64: result.pdfData,
          pdfFileName: result.pdfFileName,
          htmlContent: result.htmlContent
        }
      });
      
    } catch (excelError) {
      logger.error('Excel processing failed', {
        userId: req.user._id,
        templateId,
        error: excelError.message,
        stack: excelError.stack,
        operation: 'applyFormToTemplate'
      });
      
      // Return error instead of falling back to JSON approach
      return res.status(500).json({
        success: false,
        error: 'Excel processing failed',
        message: excelError.message,
        approach: 'EXCEL_ONLY', // No fallback
        details: 'Excel generation service encountered an error. Please check your form data and try again.'
      });
    }
  } catch (error) {
    logger.error('Error applying form data', {
      userId: req.user._id,
      templateId,
      error: error.message,
      stack: error.stack,
      operation: 'applyFormToTemplate'
    });
    res.status(500).json({ error: error.message });
  }
});

// Apply final edits (sheet updates) and return recalculated sheets (Python engine)
router.post('/templates/:templateId/apply-final', verifyToken, async (req, res) => {
  try {
    const { templateId } = req.params;
    const { updates, recalculate } = req.body || {};

    if (!Array.isArray(updates) || updates.length === 0) {
      return res.status(400).json({ success: false, error: 'No updates provided' });
    }

    let stagingRecord = null;
    let existingExcelBuffer = null;
    try {
      stagingRecord = await ReportStaging.findOne({
        user_id: req.user._id,
        template_id: templateId,
        status: 'active'
      }).sort({ createdAt: -1 });

      if (stagingRecord?.excel_data) {
        existingExcelBuffer = stagingRecord.excel_data;
      }
    } catch (stagingError) {
      logger.warn('Failed to read staging Excel during apply-final', {
        userId: req.user._id,
        templateId,
        error: stagingError.message,
        operation: 'applyFinalEdits.loadStaging'
      });
    }

    const result = await excelCalculationService.applyUpdatesAndCalculate(templateId, {
      updates,
      recalculate,
      existingExcelBuffer
    });

    if (result?.excelData) {
      try {
        const updatedBuffer = Buffer.from(result.excelData, 'base64');
        if (stagingRecord) {
          stagingRecord.excel_data = updatedBuffer;
          if (result.fileName) {
            stagingRecord.file_name = result.fileName;
          }
          await stagingRecord.save();
        } else {
          const newRecord = new ReportStaging({
            user_id: req.user._id,
            template_id: templateId,
            excel_data: updatedBuffer,
            file_name: result.fileName || `${templateId}-final.xlsx`
          });
          await newRecord.save();
        }
      } catch (stagingSaveError) {
        logger.warn('Failed to persist recalculated Excel to staging', {
          userId: req.user._id,
          templateId,
          error: stagingSaveError.message,
          operation: 'applyFinalEdits.saveStaging'
        });
      }
    }

    res.json({
      success: true,
      message: 'Final edits applied successfully',
      data: {
        fileName: result.fileName,
        excelBase64: result.excelData,
        htmlContent: result.htmlContent,
        htmlJsonData: result.htmlJsonData,
        pdfBase64: result.pdfData,
        pdfFileName: result.pdfFileName
      }
    });
  } catch (error) {
    logger.error('Error applying final edits', {
      userId: req.user._id,
      templateId,
      updatesCount: req.body?.updates?.length || 0,
      error: error.message,
      stack: error.stack,
      operation: 'applyFinalEdits'
    });
    res.status(500).json({ success: false, error: error.message || 'Failed to apply final edits' });
  }
});

router.post('/', verifyToken, async (req, res) => {
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
    
    // Log payload size for debugging
    const payloadSize = JSON.stringify(req.body).length;
    logger.debug('Report creation payload received', {
      userId: req.user._id,
      templateId,
      payloadSizeKB: (payloadSize / 1024).toFixed(2),
      operation: 'createReport'
    });
    
    // Note: Payment is now handled separately via /create-payment-order and /verify-payment endpoints
    // This endpoint creates draft reports without requiring upfront payment
    
    // Generate the Excel data
    let excelData = null;
    let jsonData = null;
    if (form_data) {
      try {
        const excelResult = await excelCalculationService.applyFormDataAndCalculate(templateId, form_data);
        if (excelResult.excelData) {
          excelData = Buffer.from(excelResult.excelData, 'base64');
        }
        if (excelResult.jsonData) {
          jsonData = excelResult.jsonData;
        }
      } catch (excelError) {
        logger.error('Error generating Excel during report creation', {
          userId: req.user._id,
          templateId,
          error: excelError.message,
          stack: excelError.stack,
          operation: 'createReport'
        });
        // Continue without Excel data for now
      }
    }
    
    const report = new Report({
      user_id: req.user._id,
      title,
      templateId,
      excel_file_url,
      excel_data: excelData, // Store Excel buffer in DB
      json_data: jsonData, // Store JSON data for browser display
      pdf_file_url,
      form_data, // Only store user input data (small)
      hidden_sheets: hidden_sheets || [],
      locked_sheets: locked_sheets || [],
      status: 'completed'
    });
    
    await report.save();
    
    // Note: The final Excel JSON will be uploaded separately via the upload-json endpoint
    // This keeps the initial report creation lightweight
    
    logger.business('Report created successfully', {
      userId: req.user._id,
      reportId: report._id,
      title,
      templateId,
      operation: 'createReport'
    });
    
    // Return minimal response
    res.status(201).json({
      _id: report._id,
      title: report.title,
      templateId: report.templateId,
      status: report.status,
      json_file_url: report.json_file_url,
      excel_download_url: `/api/reports/${report._id}/download-excel`,
      hidden_sheets: report.hidden_sheets,
      locked_sheets: report.locked_sheets,
      createdAt: report.createdAt
    });
  } catch (error) {
    logger.error('Error creating report', {
      userId: req.user._id,
      title: req.body.title,
      templateId: req.body.templateId,
      error: error.message,
      stack: error.stack,
      operation: 'createReport'
    });
    res.status(400).json({ error: error.message });
  }
});

router.get('/', verifyToken, async (req, res) => {
  try {
    let query;
    
    // Admin can see all reports
    if (req.user.role === 'admin' || req.user.role === 'super_admin') {
      query = {};
    } else {
      // Regular users can only see their own APPROVED reports
      query = { 
        user_id: req.user._id,
        validation_status: 'approved'
      };
    }
    
    const reports = await Report.find(query)
      .populate('user_id', 'name email')
      .select('-excel_data -json_data') // Don't send large binary/JSON data in list
      .sort({ createdAt: -1 });
    
    res.json({
      success: true,
      data: reports,
      message: `Found ${reports.length} reports`
    });
  } catch (error) {
    res.status(500).json({ 
      success: false,
      error: error.message,
      message: 'Failed to retrieve reports'
    });
  }
});

router.get('/:reportId', verifyToken, async (req, res) => {
  logger.debug('Report retrieval requested', {
    reportId: req.params.reportId,
    userId: req.user?._id,
    operation: 'getReportById'
  });
  try {
    const { reportId } = req.params;
    const report = await Report.findOne({ _id: reportId }).populate('user_id', 'name email');
    
    if (!report) {
      return res.status(404).json({ error: 'Report not found' });
    }
    
    // Check access permissions
    const isAdmin = req.user.role === 'admin' || req.user.role === 'super_admin';
    const isOwner = report.user_id._id.toString() === req.user._id.toString();
    
    if (!isAdmin && !isOwner) {
      return res.status(403).json({ error: 'Access denied' });
    }
    
    // Regular users can only access approved reports
    if (!isAdmin && report.validation_status !== 'approved') {
      return res.status(403).json({ 
        error: 'Report is not yet approved',
        message: 'This report is pending validation. You will be notified once it is approved.'
      });
    }
    
    res.json(report);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Upload final Excel JSON for a report (no authentication required)
router.post('/:reportId/upload-json', async (req, res) => {
  try {
    const { reportId } = req.params;
    const { finalData } = req.body;

    // Allow any report to be updated without authentication
    const report = await Report.findOne({ _id: reportId });

    if (!report) {
      return res.status(404).json({ error: 'Report not found' });
    }

    if (!finalData) {
      return res.status(400).json({ error: 'No final data provided' });
    }

    // Convert Luckysheet JSON to cell updates format for Python engine
    const updates = [];
    if (Array.isArray(finalData)) {
      finalData.forEach(sheet => {
        if (sheet.data && Array.isArray(sheet.data)) {
          const sheetName = sheet.name || 'Sheet1';
          sheet.data.forEach((row, rowIndex) => {
            if (Array.isArray(row)) {
              row.forEach((cell, colIndex) => {
                if (cell && cell.v !== undefined && cell.v !== null && cell.v !== '') {
                  // Convert column index to letter (0 = A, 1 = B, etc.)
                  const colLetter = String.fromCharCode(65 + colIndex);
                  const cellRef = `${colLetter}${rowIndex + 1}`;
                  updates.push({
                    sheet: sheetName,
                    cell: cellRef,
                    value: cell.v
                  });
                }
              });
            }
          });
        }
      });
    }

    logger.debug('Converted Luckysheet data to cell updates', {
      reportId,
      updatesCount: updates.length,
      operation: 'uploadReportJson'
    });

    // Regenerate Excel with the final updates
    let newExcelData = null;
    let newJsonData = null;
    try {
      const excelResult = await excelCalculationService.applyUpdatesAndCalculate(report.templateId, {
        updates,
        recalculate: false  // Let Excel handle automatic calculation
      });

      if (excelResult.excelData) {
        newExcelData = Buffer.from(excelResult.excelData, 'base64');
      }
      if (excelResult.jsonData) {
        newJsonData = excelResult.jsonData;
      }
    } catch (excelError) {
      logger.error('Error regenerating Excel during JSON upload', {
        reportId,
        templateId: report.templateId,
        updatesCount: updates.length,
        error: excelError.message,
        stack: excelError.stack,
        operation: 'uploadReportJson'
      });
      // Continue with JSON update only
    }

    // Log original size
    const originalJsonString = JSON.stringify(finalData, null, 2);
    const originalSize = Buffer.byteLength(originalJsonString, 'utf8');
    logger.debug('JSON upload processing started', {
      reportId,
      originalSizeKB: (originalSize / 1024).toFixed(2),
      operation: 'uploadReportJson'
    });

    // Save the final Excel JSON with compression
    try {
      const jsonFilePath = path.join(__dirname, '../uploads', `${reportId}.json`);
      const compressedJsonPath = path.join(__dirname, '../uploads', `${reportId}.json.gz`);

      // Save uncompressed version for backward compatibility
      await fs.writeFile(jsonFilePath, originalJsonString);

      // Compress and save gzipped version
      const compressed = await new Promise((resolve, reject) => {
        zlib.gzip(originalJsonString, (err, result) => {
          if (err) reject(err);
          else resolve(result);
        });
      });

      await fs.writeFile(compressedJsonPath, compressed);

      const compressedSize = compressed.length;
      const compressionRatio = ((originalSize - compressedSize) / originalSize * 100).toFixed(2);

      logger.business('JSON compression completed', {
        reportId,
        originalSizeKB: (originalSize / 1024).toFixed(2),
        compressedSizeKB: (compressedSize / 1024).toFixed(2),
        compressionRatio: `${compressionRatio}%`,
        operation: 'uploadReportJson'
      });

      // Update report with both file URLs, JSON data, and recalculated Excel data
      report.json_file_url = `/uploads/${reportId}.json`;
      report.compressed_json_url = `/uploads/${reportId}.json.gz`;
      report.json_data = finalData; // Store Luckysheet JSON data in DB for quick access
      if (newExcelData) {
        report.excel_data = newExcelData; // Store recalculated Excel binary
      }
      await report.save();

      logger.business('Report JSON files saved successfully', {
        reportId,
        jsonFileUrl: report.json_file_url,
        compressedFileUrl: report.compressed_json_url,
        excelDataSize: newExcelData ? `${newExcelData.length} bytes` : null,
        operation: 'uploadReportJson'
      });

      res.json({
        success: true,
        json_file_url: report.json_file_url,
        compressed_json_url: report.compressed_json_url,
        excel_updated: !!newExcelData,
        compression_stats: {
          original_size_kb: (originalSize / 1024).toFixed(2),
          compressed_size_kb: (compressedSize / 1024).toFixed(2),
          compression_ratio: `${compressionRatio}%`
        },
        message: 'Final Excel JSON uploaded, compressed, and Excel recalculated successfully'
      });
      
      // Update report with both file URLs and JSON data
      report.json_file_url = `/uploads/${reportId}.json`;
      report.compressed_json_url = `/uploads/${reportId}.json.gz`;
      report.json_data = finalData; // Store JSON data in DB for quick access
      await report.save();
      
      res.json({
        success: true,
        json_file_url: report.json_file_url,
        compressed_json_url: report.compressed_json_url,
        compression_stats: {
          original_size_kb: (originalSize / 1024).toFixed(2),
          compressed_size_kb: (compressedSize / 1024).toFixed(2),
          compression_ratio: `${compressionRatio}%`
        },
        message: 'Final Excel JSON uploaded and compressed successfully'
      });
      
    } catch (fileError) {
      logger.error('Error saving final JSON file', {
        reportId,
        error: fileError.message,
        stack: fileError.stack,
        operation: 'uploadReportJson'
      });
      res.status(500).json({ error: 'Failed to save Excel JSON file' });
    }
    
  } catch (error) {
    logger.error('Error uploading final JSON', {
      reportId: req.params.reportId,
      error: error.message,
      stack: error.stack,
      operation: 'uploadReportJson'
    });
    res.status(500).json({ error: error.message });
  }
});

// Serve compressed JSON files
router.get('/:reportId/download-compressed', async (req, res) => {
  try {
    const { reportId } = req.params;
    const compressedJsonPath = path.join(__dirname, '../uploads', `${reportId}.json.gz`);
    
    try {
      // Read compressed file
      const compressedData = await fs.readFile(compressedJsonPath);
      
      // Set appropriate headers for compressed content
      res.setHeader('Content-Type', 'application/json');
      res.setHeader('Content-Encoding', 'gzip');
      res.setHeader('Content-Disposition', `attachment; filename="${reportId}.json.gz"`);
      
      res.send(compressedData);
    } catch (fileError) {
      res.status(404).json({ error: 'Compressed JSON file not found' });
    }
  } catch (error) {
    logger.error('Error serving compressed JSON', {
      reportId: req.params.reportId,
      error: error.message,
      stack: error.stack,
      operation: 'downloadCompressedJson'
    });
    res.status(500).json({ error: error.message });
  }
});

// Decompress and serve JSON (for clients that can't handle gzip)
router.get('/:reportId/download-decompressed', async (req, res) => {
  try {
    const { reportId } = req.params;
    const compressedJsonPath = path.join(__dirname, '../uploads', `${reportId}.json.gz`);
    
    try {
      // Read and decompress file
      const compressedData = await fs.readFile(compressedJsonPath);
      
      const decompressed = await new Promise((resolve, reject) => {
        zlib.gunzip(compressedData, (err, result) => {
          if (err) reject(err);
          else resolve(result);
        });
      });
      
      res.setHeader('Content-Type', 'application/json');
      res.setHeader('Content-Disposition', `attachment; filename="${reportId}.json"`);
      
      res.send(decompressed);
    } catch (fileError) {
      // Fallback to uncompressed version
      const jsonFilePath = path.join(__dirname, '../uploads', `${reportId}.json`);
      try {
        const jsonData = await fs.readFile(jsonFilePath, 'utf8');
        res.setHeader('Content-Type', 'application/json');
        res.setHeader('Content-Disposition', `attachment; filename="${reportId}.json"`);
        res.send(jsonData);
      } catch (fallbackError) {
        res.status(404).json({ error: 'JSON file not found' });
      }
    }
  } catch (error) {
    logger.error('Error serving decompressed JSON', {
      reportId: req.params.reportId,
      error: error.message,
      stack: error.stack,
      operation: 'downloadDecompressedJson'
    });
    res.status(500).json({ error: error.message });
  }
});

// Download Excel file from R2 cloud storage or database (legacy)
router.get('/:reportId/download-excel', async (req, res) => {
  try {
    const { reportId } = req.params;
    
    const report = await Report.findById(reportId);
    if (!report) {
      return res.status(404).json({ error: 'Report not found' });
    }
    
    const fileName = `${reportId}_${report.templateId}.xlsx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `inline; filename="${fileName}"`);
    
    // If excel_file_url is R2 URL (cloud storage)
    if (report.excel_file_url && report.excel_file_url.includes('r2.cloudflarestorage.com')) {
      const r2Service = require('../services/cloudflareR2Service');
      const key = r2Service.extractKeyFromUrl(report.excel_file_url);
      
      if (key) {
        const fileBuffer = await r2Service.downloadFile(key);
        return res.send(fileBuffer);
      }
    }
    
    // Fallback: Excel data stored in database (legacy)
    if (report.excel_data) {
      return res.send(report.excel_data);
    }
    
    return res.status(404).json({ 
      error: 'Excel data not found',
      message: 'Please regenerate this report with the new cloud storage system'
    });
  } catch (error) {
    logger.error('Error downloading Excel', {
      reportId: req.params.reportId,
      error: error.message,
      stack: error.stack,
      operation: 'downloadExcel'
    });
    res.status(500).json({ error: error.message });
  }
});

// Get JSON data for browser display
router.get('/:reportId/json-data', async (req, res) => {
  try {
    const { reportId } = req.params;
    
    const report = await Report.findById(reportId);
    if (!report) {
      return res.status(404).json({ error: 'Report not found' });
    }
    
    if (!report.json_data) {
      return res.status(404).json({ error: 'JSON data not found for this report' });
    }
    
    res.json(report.json_data);
  } catch (error) {
    logger.error('Error getting JSON data', {
      reportId: req.params.reportId,
      error: error.message,
      stack: error.stack,
      operation: 'getJsonData'
    });
    res.status(500).json({ error: error.message });
  }
});

/**
 * @route   GET /api/reports/:reportId/finalworkings-sheet
 * @desc    Get only the FinalWorkings sheet data for display in frontend
 * @access  Public
 */
router.get('/:reportId/finalworkings-sheet', async (req, res) => {
  try {
    const { reportId } = req.params;
    
    const report = await Report.findById(reportId);
    if (!report) {
      return res.status(404).json({ success: false, error: 'Report not found' });
    }
    
    // Extract FinalWorkings sheet from stored json_data
    const finalWorkingsSheet = report.json_data?.find(sheet => sheet.name === 'FinalWorkings');
    
    if (!finalWorkingsSheet) {
      return res.status(404).json({ success: false, error: 'FinalWorkings sheet not found' });
    }
    
    res.json({
      success: true,
      data: {
        sheet: finalWorkingsSheet,
        reportId,
        templateName: report.template_name
      }
    });
  } catch (error) {
    logger.error('Error getting FinalWorkings sheet', {
      reportId: req.params.reportId,
      error: error.message,
      stack: error.stack,
      operation: 'getFinalWorkingsSheet'
    });
    res.status(500).json({ success: false, error: error.message });
  }
});

/**
 * @route   POST /api/reports/create-payment-order
 * @desc    Create Razorpay order for report payment (before final generation)
 * @access  Private
 */
router.post('/create-payment-order', verifyToken, async (req, res) => {
  try {
    const { template_id, report_title, stage_id, amount: clientAmount, selected_sheets } = req.body;
    
    if (!template_id) {
      return res.status(400).json({ success: false, error: 'template_id is required' });
    }
    
    console.log(`💰 [create-payment-order] Request for template: "${template_id}"`, {
      clientAmount,
      selectedSheetsCount: selected_sheets?.length || 0
    });
    
    let amount;
    
    // Use amount from client if provided (user selected custom sheets)
    if (clientAmount && clientAmount > 0) {
      amount = clientAmount;
      console.log(`✅ [create-payment-order] Using client-calculated amount: ₹${amount}`);
      if (selected_sheets && selected_sheets.length > 0) {
        console.log(`📋 [create-payment-order] Selected sheets:`, selected_sheets.map(s => s.sheet_name));
      }
    } else {
      // Fallback: Get pricing from template configuration
      const TemplateConfig = require('../models/TemplateConfig');
      const lookupId = template_id.toUpperCase();
      console.log(`🔍 [create-payment-order] Looking up template pricing for: "${lookupId}"`);
      
      const template = await TemplateConfig.findOne({ template_id: lookupId, is_active: true });
      
      if (template && template.pricing) {
        amount = template.pricing.effective_price || template.pricing.total_price || template.pricing.base_price || 100;
        console.log(`✅ [create-payment-order] Found template pricing: ₹${amount}`);
      } else {
        amount = 100; // Default fallback
        console.log(`⚠️ [create-payment-order] No template pricing found, using default ₹${amount}`);
      }
    }
    
    console.log(`💵 [create-payment-order] Final amount: ₹${amount}`);
    
    // Create draft report entry
    const report = new Report({
      user_id: req.user._id,
      title: report_title || `Report - ${template_id}`,
      templateId: template_id,
      report_type: template_id,
      payment: {
        status: 'pending',
        amount: amount,
        currency: 'INR'
      },
      validation_status: 'pending_payment',
      status: 'draft'
    });
    
    await report.save();
    
    // Create Razorpay order if available
    let razorpayOrder = null;
    const Razorpay = require('razorpay');
    const razorpayEnabled = process.env.RAZORPAY_KEY_ID && process.env.RAZORPAY_KEY_SECRET;
    
    if (razorpayEnabled) {
      const razorpay = new Razorpay({
        key_id: process.env.RAZORPAY_KEY_ID,
        key_secret: process.env.RAZORPAY_KEY_SECRET
      });
      
      razorpayOrder = await razorpay.orders.create({
        amount: amount * 100, // Convert to paise
        currency: 'INR',
        receipt: `report_${report._id}`,
        notes: {
          report_id: report._id.toString(),
          template_id: template_id,
          user_id: req.user._id.toString()
        }
      });
      
      report.payment.razorpay_order_id = razorpayOrder.id;
      await report.save();
    }
    
    logger.business('Payment order created for report', {
      userId: req.user._id,
      reportId: report._id,
      template_id,
      amount,
      razorpayEnabled,
      operation: 'createReportPaymentOrder'
    });
    
    res.json({
      success: true,
      data: {
        report_id: report._id,
        amount,
        currency: 'INR',
        razorpay_order_id: razorpayOrder?.id,
        razorpay_key_id: process.env.RAZORPAY_KEY_ID,
        template_id,
        report_title: report.title
      }
    });
    
  } catch (error) {
    logger.error('Error creating payment order', {
      userId: req.user._id,
      error: error.message,
      stack: error.stack,
      operation: 'createReportPaymentOrder'
    });
    res.status(500).json({ success: false, error: error.message });
  }
});

/**
 * @route   POST /api/reports/:reportId/verify-payment
 * @desc    Verify Razorpay payment and update report status
 * @access  Private
 */
router.post('/:reportId/verify-payment', verifyToken, async (req, res) => {
  try {
    const { reportId } = req.params;
    const { razorpay_payment_id, razorpay_order_id, razorpay_signature } = req.body;
    
    const report = await Report.findById(reportId);
    
    if (!report) {
      return res.status(404).json({ success: false, error: 'Report not found' });
    }
    
    if (report.user_id.toString() !== req.user._id.toString()) {
      return res.status(403).json({ success: false, error: 'Unauthorized' });
    }
    
    // Verify Razorpay signature
    const crypto = require('crypto');
    const generatedSignature = crypto
      .createHmac('sha256', process.env.RAZORPAY_KEY_SECRET || '')
      .update(`${razorpay_order_id}|${razorpay_payment_id}`)
      .digest('hex');
    
    if (generatedSignature !== razorpay_signature) {
      return res.status(400).json({ success: false, error: 'Invalid payment signature' });
    }
    
    // Update report payment status
    report.payment.status = 'completed';
    report.payment.razorpay_payment_id = razorpay_payment_id;
    report.payment.razorpay_signature = razorpay_signature;
    report.payment.paid_at = new Date();
    report.validation_status = 'pending_validation'; // Move to next stage
    await report.save();
    
    // Create order record for tracking
    const Order = require('../models/Order');
    const order = new Order({
      user_id: req.user._id,
      report_id: report._id,
      template_id: report.templateId,
      report_title: report.title,
      pack_type: 'report',
      credits: 1,
      amount_paid: report.payment.amount,
      currency: report.payment.currency,
      status: 'paid',
      payment_info: {
        rzp_order_id: razorpay_order_id,
        rzp_payment_id: razorpay_payment_id,
        rzp_signature: razorpay_signature,
        captured: true
      }
    });
    await order.save();
    
    logger.business('Payment verified successfully', {
      userId: req.user._id,
      reportId: report._id,
      orderId: order._id,
      amount: report.payment.amount,
      operation: 'verifyReportPayment'
    });
    
    res.json({
      success: true,
      message: 'Payment verified successfully',
      data: {
        report_id: report._id,
        payment_status: report.payment.status,
        order_id: order._id
      }
    });
    
  } catch (error) {
    logger.error('Error verifying payment', {
      userId: req.user._id,
      reportId: req.params.reportId,
      error: error.message,
      stack: error.stack,
      operation: 'verifyReportPayment'
    });
    res.status(500).json({ success: false, error: error.message });
  }
});

/**
 * @route   POST /api/reports/templates/:templateId/download-full-report
 * @desc    Generate full AI-enhanced report with Excel sheets (requires payment)
 * @access  Private
 */
router.post('/templates/:templateId/download-full-report', verifyToken, reportController.downloadFullReport);

/**
 * @route   GET /api/reports/download/:fileId/:type
 * @desc    Download Excel file or AI report from MongoDB
 * @access  Public (ownership verified in controller)
 */
router.get('/download/:fileId/:type', reportController.downloadFile);

module.exports = router;
