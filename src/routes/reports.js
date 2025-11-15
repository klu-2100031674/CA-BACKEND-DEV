const express = require('express');
const fs = require('fs').promises;
const path = require('path');
const zlib = require('zlib');
const Report = require('../models/Report');
const Wallet = require('../models/Wallet');
const { verifyToken } = require('../middleware/auth');
const { alterTemplateJson } = require('../services/jsonAlterService');
const excelCalculationService = require('../services/excelCalculationService');
const router = express.Router();

let templatesCache = null;
let cacheTimestamp = null;
const CACHE_DURATION = 5 * 60 * 1000; // 5 minutes

async function loadTemplates() {
  if (templatesCache && cacheTimestamp && (Date.now() - cacheTimestamp < CACHE_DURATION)) {
    return templatesCache;
  }
  
  try {
    const metaPath = path.join(__dirname, '../../templates/meta.json');
    console.log('📁 Loading templates from:', metaPath);
    const metaData = await fs.readFile(metaPath, 'utf8');
    templatesCache = JSON.parse(metaData);
    cacheTimestamp = Date.now();
    console.log('✅ Templates loaded:', templatesCache.length, 'templates');
    return templatesCache;
  } catch (error) {
    console.error('❌ Failed to load templates:', error.message, 'Path:', path.join(__dirname, '../../templates/meta.json'));
    throw new Error('Failed to load templates: ' + error.message);
  }
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
      console.log('📝 Loading template form from:', formPath);
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
      console.error('❌ Template form file error:', fileError.message);
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
    
    console.log(`🚀 [EXCEL APPROACH] Processing template: ${templateId}`);
    console.log(`📊 [EXCEL APPROACH] Form data keys: ${Object.keys(formData).length}`);
    
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
      
      console.log(`📋 [EXCEL APPROACH] Step 1: Received form data from frontend`);
      console.log(`📋 [EXCEL APPROACH] Raw payload:`, JSON.stringify(formData, null, 2).substring(0, 500) + '...');
      
      console.log(`📋 [EXCEL APPROACH] Step 2: Sending FULL payload to Excel Calculation Service (including additionalData)...`);
      
      // ⚠️ IMPORTANT: Pass the ORIGINAL formData to the service
      // The service's extractFormData() will handle unwrapping for cell data
      // The service's extractFixedAssetsSchedule() needs the full payload with additionalData
      const result = await excelCalculationService.applyFormDataAndCalculate(templateId, formData);

      console.log(`✅ [EXCEL APPROACH] Step 3: Excel recalculated successfully`);
      console.log(`✅ [EXCEL APPROACH] Step 4: Excel file generated and encoded to base64`);

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
      console.error('❌ [EXCEL APPROACH] Excel processing failed:', excelError);
      
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
    console.error('❌ [GENERAL] Error applying form data:', error);
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

    const result = await excelCalculationService.applyUpdatesAndCalculate(templateId, { updates, recalculate });

    res.json({
      success: true,
      message: 'Final edits applied successfully',
      data: {
        fileName: result.fileName,
        excelBase64: result.excelData
      }
    });
  } catch (error) {
    console.error('❌ [EXCEL APPROACH] apply-final failed:', error);
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
    console.log(`Received payload size: ${(payloadSize / 1024).toFixed(2)} KB (form data only)`);
    
    const wallet = await Wallet.findOne({ user_id: req.user._id });
    if (!wallet || wallet.report_credits < 1) {
      return res.status(400).json({ error: 'Insufficient report credits', error_code: 'INSUFFICIENT_REPORT_CREDITS'});
    }
    
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
        console.error('Error generating Excel:', excelError);
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
    
    wallet.report_credits -= 1;
    await wallet.save();
    
    console.log(`Report created successfully: ${report._id}`);
    
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
    console.error('Error creating report:', error);
    res.status(400).json({ error: error.message });
  }
});

router.get('/', verifyToken, async (req, res) => {
  try {
    const query = req.user.role === 'admin' ? {} : { user_id: req.user._id };
    
    const reports = await Report.find(query)
      .populate('user_id', 'name email')
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

router.get('/:reportId', async (req, res) => {
  try {
    const { reportId } = req.params;
    // Allow access to any report without authentication
    const report = await Report.findOne({ _id: reportId }).populate('user_id', 'name email');
    
    if (!report) {
      return res.status(404).json({ error: 'Report not found' });
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

    console.log(`Converted ${updates.length} cell updates from Luckysheet data`);

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
      console.error('Error regenerating Excel:', excelError);
      // Continue with JSON update only
    }

    // Log original size
    const originalJsonString = JSON.stringify(finalData, null, 2);
    const originalSize = Buffer.byteLength(originalJsonString, 'utf8');
    console.log(`Original JSON size: ${(originalSize / 1024).toFixed(2)} KB`);

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

      console.log(`Compressed JSON size: ${(compressedSize / 1024).toFixed(2)} KB`);
      console.log(`Compression ratio: ${compressionRatio}% reduction`);

      // Update report with both file URLs, JSON data, and recalculated Excel data
      report.json_file_url = `/uploads/${reportId}.json`;
      report.compressed_json_url = `/uploads/${reportId}.json.gz`;
      report.json_data = finalData; // Store Luckysheet JSON data in DB for quick access
      if (newExcelData) {
        report.excel_data = newExcelData; // Store recalculated Excel binary
      }
      await report.save();

      console.log(`Final Excel JSON saved to: ${report.json_file_url}`);
      console.log(`Compressed version saved to: ${report.compressed_json_url}`);
      if (newExcelData) {
        console.log(`Recalculated Excel binary stored in database (${newExcelData.length} bytes)`);
      }

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
      
      console.log(`Final Excel JSON saved to: ${report.json_file_url}`);
      console.log(`Compressed version saved to: ${report.compressed_json_url}`);
      
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
      console.error('Error saving final JSON file:', fileError);
      res.status(500).json({ error: 'Failed to save Excel JSON file' });
    }
    
  } catch (error) {
    console.error('Error uploading final JSON:', error);
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
    console.error('Error serving compressed JSON:', error);
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
    console.error('Error serving decompressed JSON:', error);
    res.status(500).json({ error: error.message });
  }
});

// Download Excel file from database
router.get('/:reportId/download-excel', async (req, res) => {
  try {
    const { reportId } = req.params;
    
    const report = await Report.findById(reportId);
    if (!report) {
      return res.status(404).json({ error: 'Report not found' });
    }
    
    if (!report.excel_data) {
      return res.status(404).json({ error: 'Excel data not found for this report' });
    }
    
    // Set headers for Excel display in browser
    const fileName = `${reportId}_${report.templateId}.xlsx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `inline; filename="${fileName}"`);
    
    // Send the Excel buffer
    res.send(report.excel_data);
  } catch (error) {
    console.error('Error downloading Excel:', error);
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
    console.error('Error getting JSON data:', error);
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
    console.error('Error getting FinalWorkings sheet:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

/**
 * @route   POST /api/reports/templates/:templateId/download-full-report
 * @desc    Generate full AI-enhanced report with Excel sheets
 * @access  Public (optional authentication)
 */
router.post('/templates/:templateId/download-full-report', async (req, res) => {
  try {
    const { templateId } = req.params;
    const formData = req.body;
    const geminiApiKey = req.body.geminiApiKey || process.env.GEMINI_API_KEY;

    console.log('🚀 [Full Report] Starting generation...');
    console.log(`  Template: ${templateId}`);
    console.log(`  User: ${req.user ? req.user._id : 'Not authenticated'}`);

    if (!geminiApiKey) {
      return res.status(400).json({
        success: false,
        error: 'Gemini API key is required. Provide it in request body or set GEMINI_API_KEY environment variable.'
      });
    }

    // Call the service to generate full report
    const result = await excelCalculationService.generateFullReport(
      templateId,
      formData,
      geminiApiKey
    );

    if (!result.success) {
      return res.status(500).json({
        success: false,
        error: result.error || 'Failed to generate full report'
      });
    }

    console.log('✅ [Full Report] Generation complete');
    console.log(`  File: ${result.fullReportFileName}`);
    console.log(`  Size: ${result.fullReportBase64?.length || 0} bytes (base64)`);

    res.json({
      success: true,
      excelData: result.excelData,
      pdfData: result.pdfData,
      fullReportData: result.fullReportData,
      fullReportFileName: result.fullReportFileName,
      aiInsights: result.aiInsights,
      data: {
        fullReportFileName: result.fullReportFileName,
        fullReportBase64: result.fullReportData,
        excelFileName: result.fileName || `frcc1_report_${Date.now()}.xlsx`,
        excelBase64: result.excelData,
        pdfFileName: result.pdfFileName,
        pdfBase64: result.pdfData,
        aiInsights: result.aiInsights
      }
    });
  } catch (error) {
    console.error('❌ [Full Report] Generation failed:', error);

    // Handle template validation errors
    if (error.message && error.message.includes('Template') && error.message.includes('not found')) {
      return res.status(404).json({
        success: false,
        error: 'Template not found'
      });
    }

    res.status(500).json({
      success: false,
      error: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
});

module.exports = router;