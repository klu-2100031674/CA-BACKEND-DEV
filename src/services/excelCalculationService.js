const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const templateMappingService = require('./templateMappingService');

/**
 * Excel Calculation Service (Python-Powered)
 * ------------------------------------------
 * - Invokes a Python script to perform Excel calculations.
 * - Loads Excel template.
 * - Updates cells using form data.
 * - Calculates all formulas.
 * - Extracts data from a specified sheet ('finalworkig' by default).
 * - Returns parsed data.
 * - Uses dynamic template mappings to prevent overwriting formulas.
 */
class ExcelCalculationService {
  constructor() {
    this.templatesPath = path.join(__dirname, '../../templates/excel');
    this.pythonEnginePath = path.join(__dirname, '../python-engine');
    // Use virtual environment Python where openai is installed
    const venvDir = process.platform === 'win32' ? 'Scripts' : 'bin';
    const pythonExe = process.platform === 'win32' ? 'python.exe' : 'python';
    const venvPythonPath = path.join(this.pythonEnginePath, '.venv', venvDir, pythonExe);
    
    // Check if virtual environment Python exists; otherwise, use system Python
    this.pythonExecutable = fs.existsSync(venvPythonPath) ? venvPythonPath : 'python';
    
    console.log(`[ExcelCalculationService] Python executable path: ${this.pythonExecutable}`);
    this.tempDir = process.env.TEMP_DIR || path.join(__dirname, '../../temp');
    console.log(`[ExcelCalculationService] Temp directory: ${this.tempDir}`);
  }

  // Extract cell mapping from various payload formats
  extractFormData(payload, templateId = null) {
    console.log('[ExcelCalculationService] Extracting form data from payload');
    console.log('[ExcelCalculationService] Template ID:', templateId);
    console.log('[ExcelCalculationService] Payload keys:', Object.keys(payload));

    // Normalize template ID
    const normalizedTemplateId = templateId ? templateMappingService.normalizeTemplateId(templateId) : null;
    console.log('[ExcelCalculationService] Normalized Template ID:', normalizedTemplateId);

    // Handle different payload formats
    let cellData = {};

    // Format 1: Direct cell mapping (from test files)
    if (payload && typeof payload === 'object' && !payload.formData) {
      // Check if it looks like cell data (has keys like 'i4', 'I4', etc.)
      const keys = Object.keys(payload);
      if (keys.some(key => key.match(/^[d-eh-j]\d+$/i))) {
        cellData = payload;
        console.log('[ExcelCalculationService] Found direct cell mapping format');
      }
    }

    // Format 2: Nested formData structure (from frontend)
    if (payload && payload.formData) {
      if (payload.formData.excelData) {
        cellData = payload.formData.excelData;
        console.log('[ExcelCalculationService] Found nested formData.excelData format');
      } else if (payload.formData.formData && payload.formData.formData.excelData) {
        cellData = payload.formData.formData.excelData;
        console.log('[ExcelCalculationService] Found deeply nested formData.formData.excelData format');
      }
    }

    // Template-specific cell validation
    const normalizedData = {};

    // Define allowed columns for each template
    const templateColumnRules = {
      'CC1': /^[b-eh-j]\d+$/,  // b,d,e for fixed assets + h,i,j for main sections
      'CC2': /^[b-eh-j]\d+$/,  // b,d,e for fixed assets + h,i,j for main sections
      'CC3': /^[b-eh-j]\d+$/,  // b,d,e for fixed assets + h,i,j for main sections
      'CC4': /^[d-eh-j]\d+$/,  // d,e for fixed assets + h,i,j for main sections
      'CC5': /^[d-eh-j]\d+$/,  // d,e for fixed assets + h,i,j for main sections
      'CC6': /^[d-eh-j]\d+$/   // d,e for fixed assets + h,i,j for main sections
    };

    const allowedPattern = templateColumnRules[normalizedTemplateId] || /^[h-j]\d+$/; // Default to h-j for unknown templates

    for (const [key, value] of Object.entries(cellData)) {
      const lowerKey = key.toLowerCase();
      if (lowerKey.match(allowedPattern)) {
        // Extract value from {label, value} object if present, otherwise use value directly
        const cellValue = (value && typeof value === 'object' && 'value' in value) ? value.value : value;
        normalizedData[lowerKey] = cellValue;
      }
    }

    console.log('[ExcelCalculationService] Extracted cell data (before filtering):', Object.keys(normalizedData).length, 'cells');

    // Apply template-specific filtering to prevent overwriting formulas
    let filteredData = normalizedData;
    if (templateId) {
      filteredData = templateMappingService.filterWritableCells(templateId, normalizedData);
      console.log('[ExcelCalculationService] After template filtering:', Object.keys(filteredData).length, 'cells');
    } else {
      console.warn('[ExcelCalculationService] No templateId provided - skipping formula protection filter');
    }

    return normalizedData; // Return normalizedData to ensure extracted values
  }

  // Extract Fixed Assets Schedule items and map to D/E columns (uses dynamic mapping)
  extractFixedAssetsSchedule(formDataPayload, templateId = 'Format CC1') {
    console.log('🔧 [ExcelCalculationService] ========================================');
    console.log(`🔧 [ExcelCalculationService] EXTRACTING FIXED ASSETS SCHEDULE - ${templateId}`);
    console.log('🔧 [ExcelCalculationService] ========================================');

    // Normalize template ID for consistent checking
    const normalizedTemplateId = templateId ? templateMappingService.normalizeTemplateId(templateId) : null;
    console.log(`🔧 [ExcelCalculationService] Normalized template ID: ${normalizedTemplateId}`);

    // CC1, CC2, and CC3 now use cell mappings instead of row mappings, so skip this extraction
    if (normalizedTemplateId === 'CC1' || normalizedTemplateId === 'CC2' || normalizedTemplateId === 'CC3') {
      console.log(`✅ [ExcelCalculationService] Skipping fixed assets extraction for ${templateId} (normalized: ${normalizedTemplateId}) (uses cell mappings)`);
      return [];
    }
    
    let fixedAssetsSchedule = null;
    
    // Try different payload structures (check deep nested path first)
    if (formDataPayload?.formData?.formData?.additionalData?.["Fixed Assets Schedule"]) {
      fixedAssetsSchedule = formDataPayload.formData.formData.additionalData["Fixed Assets Schedule"];
      console.log('✅ [ExcelCalculationService] Found Fixed Assets Schedule in formData.formData.additionalData');
    } else if (formDataPayload?.formData?.additionalData?.["Fixed Assets Schedule"]) {
      fixedAssetsSchedule = formDataPayload.formData.additionalData["Fixed Assets Schedule"];
      console.log('✅ [ExcelCalculationService] Found Fixed Assets Schedule in formData.additionalData');
    } else if (formDataPayload?.formData?.formData?.["Fixed Assets Schedule"]) {
      fixedAssetsSchedule = formDataPayload.formData.formData["Fixed Assets Schedule"];
      console.log('✅ [ExcelCalculationService] Found Fixed Assets Schedule in formData.formData');
    } else if (formDataPayload?.formData?.["Fixed Assets Schedule"]) {
      fixedAssetsSchedule = formDataPayload.formData["Fixed Assets Schedule"];
      console.log('✅ [ExcelCalculationService] Found Fixed Assets Schedule in formData');
    } else if (formDataPayload?.additionalData?.["Fixed Assets Schedule"]) {
      fixedAssetsSchedule = formDataPayload.additionalData["Fixed Assets Schedule"];
      console.log('✅ [ExcelCalculationService] Found Fixed Assets Schedule in additionalData');
    }
    
    if (!fixedAssetsSchedule) {
      console.log('❌ [ExcelCalculationService] No Fixed Assets Schedule found in payload');
      console.log('❌ [ExcelCalculationService] Payload keys:', Object.keys(formDataPayload || {}));
      return [];
    }
    
    const updates = [];
    
    // Get Fixed Assets mapping from template mapping service
    let categoryRowMapping = templateMappingService.getFixedAssetsMapping(templateId);
    
    if (!categoryRowMapping || Object.keys(categoryRowMapping).length === 0) {
      console.warn('⚠️  [ExcelCalculationService] No fixed assets mapping found for template, using defaults');
      // Fallback to CC6 mapping
      categoryRowMapping = {
        "plant_machinery": 135,
        "service_equipment": 145,
        "shed_civil": 155,
        "land": 165,
        "electrical": 168,
        "electronic": 178,
        "furniture": 188,
        "vehicles": 198,
        "other_assets": 208,
        "capital_wip": 217
      };
    }
    
    console.log('📊 [ExcelCalculationService] Using Fixed Assets Mapping:', categoryRowMapping);
    
    // Map frontend category names to backend mapping keys
    const categoryNameMap = {
      "Plant and Machinery": "plant_machinery",
      "Service Equipment": "service_equipment",
      "Civil works & Shed Construction": "shed_civil",
      "Land": "land",
      "Electrical Items & fittings": "electrical",
      "Electronic Items": "electronic",
      "Furniture and Fittings": "furniture",
      "Vehicles": "vehicles",
      "Live stock": "other_assets",
      "Other Assets": "capital_wip"
    };
    
    for (const [categoryName, categoryData] of Object.entries(fixedAssetsSchedule)) {
      if (!categoryData.items || !Array.isArray(categoryData.items)) continue;
      
      // Map frontend category name to backend key
      const mappingKey = categoryNameMap[categoryName] || categoryName.toLowerCase().replace(/\s+/g, '_');
      const startRow = categoryRowMapping[mappingKey];
      
      if (!startRow) {
        console.log(`[ExcelCalculationService] Unknown category: ${categoryName} (key: ${mappingKey}), skipping`);
        continue;
      }
      
      // Write each item in the category
      categoryData.items.forEach((item, index) => {
        const row = startRow + index;
        updates.push({ sheet: 'Assumptions.1', cell: `d${row}`, value: item.description || '' });
        updates.push({ sheet: 'Assumptions.1', cell: `e${row}`, value: item.amount || 0 });
        console.log(`[ExcelCalculationService] ${categoryName}: d${row}=${item.description}, e${row}=${item.amount}`);
      });
    }
    
    console.log(`[ExcelCalculationService] Extracted ${updates.length} fixed asset items`);
    return updates;
  }

  // ────────────────────────────────────────────────────────────────
  //  Main entry: apply data and extract JSON by calling Python script
  async applyFormDataAndCalculate(templateId, formDataPayload) {
    try {
      console.log(`[ExcelCalculationService] Starting Python calculation for template: ${templateId}`);

      // Extract cell data from the payload (with template-specific filtering)
      const cellData = this.extractFormData(formDataPayload, templateId);
      console.log(`[ExcelCalculationService] Extracted cell data: ${Object.keys(cellData).length} cells`);

      const updates = [];
      for (const [cell, value] of Object.entries(cellData)) {
        updates.push({ sheet: 'Assumptions.1', cell, value });
      }
      console.log(`[ExcelCalculationService] Added ${updates.length} cell updates`);
      
      // Extract and add Fixed Assets Schedule items (pass templateId for correct mapping)
      const fixedAssetsUpdates = this.extractFixedAssetsSchedule(formDataPayload, templateId);
      updates.push(...fixedAssetsUpdates);
      console.log(`[ExcelCalculationService] Added ${fixedAssetsUpdates.length} fixed assets updates`);
      
      const inputData = {
        updates,
        recalculate: false, // Let Excel handle automatic calculation
      };
      console.log(`[ExcelCalculationService] Total updates: ${updates.length}`);

      // Resolve template path (handle different naming conventions)
      const templatePath = this.resolveTemplatePath(templateId);
      console.log(`[ExcelCalculationService] Using template path: ${templatePath}`);
      
      const scriptPath = path.join(this.pythonEnginePath, 'excel_calculator.py');
      console.log(`[ExcelCalculationService] Using script path: ${scriptPath}`);
      
      console.log(`[ExcelCalculationService] Running Python script with executable: ${this.pythonExecutable}`);
      const result = await this.runPythonScript(scriptPath, [templatePath, JSON.stringify(inputData)]);
      
      console.log(`[ExcelCalculationService] Python script finished.`);
      const excelResult = this.transformPythonResult(JSON.parse(result));

      // PDF is now generated directly in Python, no need for separate generation
      console.log(`[ExcelCalculationService] PDF generated directly from Excel sheet`);

      return excelResult;

    } catch (error) {
      console.error(`[ExcelCalculationService] Error during Python script execution:`, error);
      throw new Error('Failed to calculate Excel data using Python engine.');
    }
  }

  // Resolve template path based on templateId (handle different naming conventions)
  resolveTemplatePath(templateId) {
    // Map template IDs to actual file names
    const templateFileMap = {
      'CC1': 'format CC1.xlsx',
      'frcc1': 'format CC1.xlsx',
      'Format CC1': 'format CC1.xlsx',
      'CC2': 'format CC2.xlsx',
      'frcc2': 'format CC2.xlsx',
      'Format CC2': 'format CC2.xlsx',
      'CC3': 'Format CC3.xlsx',
      'frcc3': 'Format CC3.xlsx',
      'Format CC3': 'Format CC3.xlsx',
      'CC4': 'format CC4.xlsx',
      'frcc4': 'format CC4.xlsx',
      'Format CC4': 'format CC4.xlsx',
      'CC5': 'format CC5.xlsx',
      'frcc5': 'format CC5.xlsx',
      'Format CC5': 'format CC5.xlsx',
      'CC6': 'format CC6.xlsx',
      'frcc6': 'format CC6.xlsx',
      'Format CC6': 'format CC6.xlsx'
    };

    const filename = templateFileMap[templateId] || `${templateId}.xlsx`;
    const templatePath = path.join(this.templatesPath, filename);
    
    // Check if file exists
    if (!fs.existsSync(templatePath)) {
      throw new Error(`Template file not found: ${filename} (templateId: ${templateId})`);
    }
    
    return templatePath;
  }

  // Apply arbitrary updates across any sheets and calculate
  async applyUpdatesAndCalculate(templateId, updatesPayload) {
    try {
      console.log(`[ExcelCalculationService] Starting Python calculation (applyUpdates) for template: ${templateId}`);

      // Expecting updatesPayload = { updates: [{sheet, cell, value}, ...], recalculate?: boolean }
      const inputData = {
        updates: Array.isArray(updatesPayload?.updates) ? updatesPayload.updates : [],
        recalculate: Boolean(updatesPayload?.recalculate ?? false), // Default to false, let Excel auto-calculate
      };

      const templatePath = this.resolveTemplatePath(templateId);
      const scriptPath = path.join(this.pythonEnginePath, 'excel_calculator.py');

      const result = await this.runPythonScript(scriptPath, [templatePath, JSON.stringify(inputData)]);
      console.log(`[ExcelCalculationService] Python script finished (applyUpdates).`);
      return this.transformPythonResult(JSON.parse(result));
    } catch (error) {
      console.error(`[ExcelCalculationService] Error during Python script execution (applyUpdates):`, error);
      throw new Error('Failed to calculate Excel data using Python engine.');
    }
  }

  // Generate full AI-enhanced report (all in Python)
  async generateFullReport(templateId, formDataPayload, apiKey, aiProvider = 'grok') {
    try {
      console.log(`[ExcelCalculationService] Starting FULL REPORT generation for template: ${templateId} using ${aiProvider}`);

      // Get API key based on provider
      let finalApiKey = apiKey;
      if (!finalApiKey) {
        if (aiProvider === 'grok') {
          finalApiKey = process.env.GROK_API_KEY;
        } else {
          finalApiKey = process.env.PERPLEXITY_API_KEY;
        }
      }

      if (!finalApiKey) {
        throw new Error(`${aiProvider.toUpperCase()} API key is required. Set ${aiProvider.toUpperCase()}_API_KEY environment variable or provide in request.`);
      }

      // Extract cell data from the payload (with template-specific filtering)
      const cellData = this.extractFormData(formDataPayload, templateId);
      
      // Convert cell data to updates array
      const updates = [];
      for (const [cell, value] of Object.entries(cellData)) {
        updates.push({ sheet: 'Assumptions.1', cell, value });
      }
      
      // Extract and add Fixed Assets Schedule items (pass templateId for correct mapping)
      const fixedAssetsUpdates = this.extractFixedAssetsSchedule(formDataPayload, templateId);
      updates.push(...fixedAssetsUpdates);

      console.log(`[ExcelCalculationService] Prepared ${updates.length} Excel updates for AI report generation`);
      console.log(`[ExcelCalculationService] - Cell updates: ${updates.length - fixedAssetsUpdates.length}`);
      console.log(`[ExcelCalculationService] - Fixed assets updates: ${fixedAssetsUpdates.length}`);

      // Build input data for Python script
      const inputData = {
        updates,  // Excel cell updates (including Fixed Assets in D/E columns)
        recalculate: false,  // Let Excel auto-calculate
        generateFullReport: true,  // Enable full report generation
      };

      // Add API key based on provider
      if (aiProvider === 'grok') {
        inputData.grokApiKey = finalApiKey;
      } else {
        inputData.perplexityApiKey = finalApiKey;
      }

      console.log('[ExcelCalculationService] Updates sample:', JSON.stringify(updates.slice(0, 3)));

      const templatePath = this.resolveTemplatePath(templateId);
      const scriptPath = path.join(this.pythonEnginePath, 'excel_calculator.py');
      
      const result = await this.runPythonScript(scriptPath, [templatePath, JSON.stringify(inputData)]);
      
      console.log(`[ExcelCalculationService] Full report generation finished.`);
      console.log(`[ExcelCalculationService] Python output (first 500 chars):`, result.substring(0, 500));
      
      let parsedResult;
      try {
        parsedResult = JSON.parse(result);
      } catch (parseError) {
        console.error('[ExcelCalculationService] Failed to parse Python output as JSON');
        console.error('[ExcelCalculationService] Parse error:', parseError.message);
        console.error('[ExcelCalculationService] Raw output:', result);
        throw new Error(`Failed to parse Python script output: ${parseError.message}`);
      }
      
      const excelResult = this.transformPythonResult(parsedResult);

      return excelResult;

    } catch (error) {
      console.error(`[ExcelCalculationService] Error during full report generation:`, error);
      console.error(`[ExcelCalculationService] Error stack:`, error.stack);
      throw error; // Rethrow the actual error instead of generic message
    }
  }

  transformPythonResult(rawResult) {
    console.log('[transformPythonResult] Raw result type:', typeof rawResult);
    console.log('[transformPythonResult] Raw result keys:', rawResult ? Object.keys(rawResult) : 'null');
    
    try {
      // Try to parse if it's a string
      let result = rawResult;
      if (typeof rawResult === 'string') {
        console.log('[transformPythonResult] Parsing JSON string...');
        result = JSON.parse(rawResult);
        console.log('[transformPythonResult] JSON parsed successfully');
      }
      
      console.log('[transformPythonResult] Success field:', result?.success);
      console.log('[transformPythonResult] Error field:', result?.error);
      
      if (!result || result.success === false) {
        const errorMessage = result?.error || 'Python engine returned an error';
        console.error('[transformPythonResult] Python script failed:', errorMessage);
        throw new Error(errorMessage);
      }

      const meta = result._meta || {};
      const verificationCopy = meta.verificationCopy ? path.normalize(meta.verificationCopy) : null;
      const verificationFileName = verificationCopy ? path.basename(verificationCopy) : null;
      const fileUrl = verificationFileName ? `/temp/${verificationFileName}` : null;

      return {
        relativePath: fileUrl,
        fileName: verificationFileName,
        excelData: result.excelData,
        jsonData: result.jsonData,
        allSheetsData: result.allSheetsData || {},
        formattedWCData: result.formattedWCData || {},
        htmlContent: result.htmlContent,
        htmlJsonData: result.htmlJsonData || {},
        pdfData: result.pdfData,
        pdfFileName: result.pdfFileName,
        meta: result._meta || {}
      };
      
    } catch (parseError) {
      console.error('[transformPythonResult] JSON parse error:', parseError.message);
      console.error('[transformPythonResult] Raw result (first 500 chars):', rawResult?.substring(0, 500));
      
      // Try to extract error from raw result if it contains error information
      if (typeof rawResult === 'string' && rawResult.includes('error')) {
        try {
          // Try to find error message in the string
          const errorMatch = rawResult.match(/"error"\s*:\s*"([^"]+)"/);
          if (errorMatch) {
            throw new Error(`Python script error: ${errorMatch[1]}`);
          }
        } catch (e) {
          // Ignore
        }
      }
      
      throw new Error(`Failed to parse Python script output: ${parseError.message}`);
    }
  }

  // ────────────────────────────────────────────────────────────────
  //  Utility to run a Python script and get its output
  runPythonScript(scriptPath, args) {
    return new Promise((resolve, reject) => {
      console.log(`[runPythonScript] Executing: ${this.pythonExecutable} ${scriptPath} ${args.join(' ')}`);
      const pythonProcess = spawn(this.pythonExecutable, [scriptPath, ...args]);

      let stdout = '';
      let stderr = '';

      pythonProcess.stdout.on('data', (data) => {
        stdout += data.toString();
      });

      pythonProcess.stderr.on('data', (data) => {
        const stderrText = data.toString();
        stderr += stderrText;
        // Log Python stderr in real-time for debugging
        console.log('[Python stderr]', stderrText);
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

  // Housekeeping: delete temp files older than N hours from backend/temp
  async cleanupTempFiles(maxAgeHours = 24) {
    const fs = require('fs').promises;
    try {
      const dir = this.tempDir;
      const entries = await fs.readdir(dir, { withFileTypes: true });
      const now = Date.now();
      const maxAgeMs = Math.max(1, Number(maxAgeHours)) * 60 * 60 * 1000;
      let deleted = 0;

      await Promise.all(
        entries.map(async (ent) => {
          if (!ent.isFile()) return;
          const filePath = path.join(dir, ent.name);
          try {
            const stat = await fs.stat(filePath);
            if (now - stat.mtimeMs > maxAgeMs) {
              await fs.unlink(filePath);
              deleted += 1;
            }
          } catch (_) {
            // ignore individual file errors
          }
        })
      );
      return deleted;
    } catch (err) {
      // If temp dir doesn't exist or other error, don't crash server
      return 0;
    }
  }
}

module.exports = new ExcelCalculationService();
