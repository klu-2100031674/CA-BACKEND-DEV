const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const fsPromises = fs.promises;
const templateMappingService = require('./templateMappingService');
const logger = require('../utils/logger');
const TemplateConfig = require('../models/TemplateConfig');

const TEMPLATE_SHEET_CONFIG = {
  TERM_LOAN_SERVICE_WITHOUT_STOCK: {
    aliasMap: {
      finalworking: 'Final workings',
      finalworkings: 'Final workings',
      finalwork: 'Final workings',
      mpbf: 'MPBF ',
      mpbfformula: 'MPBF ',
      mpbfmethod1: 'MPBF ',
      mpbfmethod2: 'MPBF ',
      workingsforsensitivity1: 'workings for sensitivity1',
      workingsforsensittivity1: 'workings for sensitivity1',
      workingsforsensitvity1: 'workings for sensitivity1',
      gaurantors: 'Gaurantors',
      guarantors: 'Gaurantors',
      bepanalysis: 'BEP analysis',
      sheetone: 'Sheet1',
      sheet1: 'Sheet1'
    }
  },
  TERM_LOAN_MANUFACTURING_SERVICE_WITH_STOCK: {
    aliasMap: {
      finalworking: 'Final workings',
      finalworkings: 'Final workings',
      finalwork: 'Final workings',
      mpbf: 'MPBF ',
      mpbfformula: 'MPBF ',
      mpbfmethod1: 'MPBF ',
      mpbfmethod2: 'MPBF ',
      workingsforsensitivity1: 'workings for sensitivity1',
      workingsforsensittivity1: 'workings for sensitivity1',
      workingsforsensitvity1: 'workings for sensitivity1',
      gaurantors: 'Gaurantors',
      guarantors: 'Gaurantors',
      bepanalysis: 'BEP analysis',
      sheetone: 'Sheet1',
      sheet1: 'Sheet1'
    }
  },
  // CC1 and CC2 use "FinalWorkings" (no space)
  CC1: {
    aliasMap: {
      finalworking: 'FinalWorkings',
      finalworkings: 'FinalWorkings',
      'final workings': 'FinalWorkings',
      'final working': 'FinalWorkings'
    }
  },
  CC2: {
    aliasMap: {
      finalworking: 'FinalWorkings',
      finalworkings: 'FinalWorkings',
      'final workings': 'FinalWorkings',
      'final working': 'FinalWorkings'
    }
  },
  // CC3 and CC4 use "Finalworkings" (lowercase w)
  CC3: {
    aliasMap: {
      finalworking: 'Finalworkings',
      finalworkings: 'Finalworkings',
      'final workings': 'Finalworkings',
      'final working': 'Finalworkings',
      'FinalWorkings': 'Finalworkings'
    }
  },
  CC4: {
    aliasMap: {
      finalworking: 'Finalworkings',
      finalworkings: 'Finalworkings',
      'final workings': 'Finalworkings',
      'final working': 'Finalworkings',
      'FinalWorkings': 'Finalworkings'
    }
  },
  // CC5 uses "FinalWorkings" (no space)
  CC5: {
    aliasMap: {
      finalworking: 'FinalWorkings',
      finalworkings: 'FinalWorkings',
      'final workings': 'FinalWorkings',
      'final working': 'FinalWorkings'
    }
  },
  // CC6 uses "Final workings" (with space)
  CC6: {
    aliasMap: {
      finalworking: 'Final workings',
      finalworkings: 'Final workings',
      'final workings': 'Final workings',
      'final working': 'Final workings'
    }
  }
};

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
    
    logger.debug('Python executable configured', {
      operation: 'constructor',
      pythonExecutable: this.pythonExecutable
    });
    this.tempDir = process.env.TEMP_DIR || path.join(__dirname, '../../temp');
    logger.debug('Temp directory configured', {
      operation: 'constructor',
      tempDir: this.tempDir
    });
  }

  // Extract cell mapping from various payload formats
  extractFormData(payload, templateId = null) {
    logger.debug('Starting form data extraction', {
      operation: 'extractFormData',
      templateId,
      payloadKeys: Object.keys(payload)
    });

    // Normalize template ID
    const normalizedTemplateId = templateId ? templateMappingService.normalizeTemplateId(templateId) : null;
    logger.debug('Template ID normalized', {
      operation: 'extractFormData',
      originalTemplateId: templateId,
      normalizedTemplateId
    });

    // Handle different payload formats
    let cellData = {};

    // Format 1: Direct cell mapping (from test files)
    if (payload && typeof payload === 'object' && !payload.formData) {
      // Check if it looks like cell data (has keys like 'i4', 'I4', etc.)
      const keys = Object.keys(payload);
      if (keys.some(key => key.match(/^[d-eh-j]\d+$/i))) {
        cellData = payload;
        logger.debug('Direct cell mapping format detected', {
          operation: 'extractFormData',
          cellKeys: keys.length
        });
      }
    }

    // Format 2: Nested formData structure (from frontend)
    if (payload && payload.formData) {
      if (payload.formData.excelData) {
        cellData = { ...payload.formData.excelData };
        logger.debug('Nested formData.excelData format detected', {
          operation: 'extractFormData'
        });
        
        // Also merge Loan Percentage Cells if present (for Term Loan forms)
        if (payload.formData['Loan Percentage Cells']) {
          Object.assign(cellData, payload.formData['Loan Percentage Cells']);
          logger.debug('Merged Loan Percentage Cells into cellData', {
            operation: 'extractFormData',
            loanCells: Object.keys(payload.formData['Loan Percentage Cells'])
          });
        }
      } else if (payload.formData.formData && payload.formData.formData.excelData) {
        cellData = { ...payload.formData.formData.excelData };
        logger.debug('Deeply nested formData.formData.excelData format detected', {
          operation: 'extractFormData'
        });
        
        // Also merge Loan Percentage Cells if present
        if (payload.formData.formData['Loan Percentage Cells']) {
          Object.assign(cellData, payload.formData.formData['Loan Percentage Cells']);
          logger.debug('Merged Loan Percentage Cells into cellData (deep)', {
            operation: 'extractFormData',
            loanCells: Object.keys(payload.formData.formData['Loan Percentage Cells'])
          });
        }
      } 
      // Format 3: Section-based structure (Term Loan form)
      else if (payload.formData) {
        // Extract all cell references from nested sections
        const extractCellsFromObject = (obj) => {
          const cells = {};
          for (const [key, value] of Object.entries(obj)) {
            // Check if key is a cell reference (e.g., i7, d118, e241, k28)
            if (typeof key === 'string' && key.match(/^[d-k]\d+$/i)) {
              cells[key.toLowerCase()] = value;
            }
            // Recursively check nested objects (sections like "General Information", "Schedule for Assets")
            else if (value && typeof value === 'object' && !Array.isArray(value)) {
              Object.assign(cells, extractCellsFromObject(value));
            }
          }
          return cells;
        };
        
        cellData = extractCellsFromObject(payload.formData);
        logger.debug('Section-based format detected, extracted cell references', {
          operation: 'extractFormData',
          cellKeys: Object.keys(cellData).length
        });
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
      'CC6': /^[d-eh-j]\d+$/,  // d,e for fixed assets + h,i,j for main sections
      'TERM_LOAN_SERVICE_WITHOUT_STOCK': /^[d-eh-jk]\d+$/,  // d,e for assets + h,i,j for main sections + k for loan percentages
      'TERM_LOAN_MANUFACTURING_SERVICE_WITH_STOCK': /^[d-eh-jk]\d+$/,  // d,e for assets + h,i,j for main sections + k for loan percentages
      'TERM_LOAN_CC': /^[d-eh-jk]\d+$/  // d,e for assets + h,i,j for main sections + k for loan percentages
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

    logger.debug('Cell data extracted and normalized', {
      operation: 'extractFormData',
      cellCount: Object.keys(normalizedData).length
    });

    // Apply template-specific filtering to prevent overwriting formulas
    let filteredData = normalizedData;
    if (templateId) {
      filteredData = templateMappingService.filterWritableCells(templateId, normalizedData);
      logger.debug('Template filtering applied', {
        operation: 'extractFormData',
        templateId,
        filteredCellCount: Object.keys(filteredData).length
      });
    } else {
      logger.warn('Template filtering skipped - no templateId provided', {
        operation: 'extractFormData'
      });
    }

    return normalizedData; // Return normalizedData to ensure extracted values
  }

  // Extract Fixed Assets Schedule items and map to D/E columns (uses dynamic mapping)
  extractFixedAssetsSchedule(formDataPayload, templateId = 'Format CC1') {
    logger.info('Starting fixed assets schedule extraction', {
      operation: 'extractFixedAssetsSchedule',
      templateId
    });

    // Normalize template ID for consistent checking
    const normalizedTemplateId = templateId ? templateMappingService.normalizeTemplateId(templateId) : null;
    logger.debug('Template ID normalized for fixed assets extraction', {
      operation: 'extractFixedAssetsSchedule',
      originalTemplateId: templateId,
      normalizedTemplateId
    });

    // CC1, CC2, CC3, CC4, CC5, CC6, and TERM_LOAN now use cell mappings instead of row mappings, so skip this extraction
    if (normalizedTemplateId === 'CC1' || normalizedTemplateId === 'CC2' || normalizedTemplateId === 'CC3' || normalizedTemplateId === 'CC4' || normalizedTemplateId === 'CC5' || normalizedTemplateId === 'CC6' || normalizedTemplateId === 'TERM_LOAN_SERVICE_WITHOUT_STOCK' || normalizedTemplateId === 'TERM_LOAN_MANUFACTURING_SERVICE_WITH_STOCK') {
      logger.info('Fixed assets extraction skipped - using cell mappings', {
        operation: 'extractFixedAssetsSchedule',
        templateId,
        normalizedTemplateId
      });
      return [];
    }
    
    let fixedAssetsSchedule = null;
    
    // Try different payload structures (check deep nested path first)
    if (formDataPayload?.formData?.formData?.additionalData?.["Fixed Assets Schedule"]) {
      fixedAssetsSchedule = formDataPayload.formData.formData.additionalData["Fixed Assets Schedule"];
      logger.debug('Fixed assets schedule found in deep nested path', {
        operation: 'extractFixedAssetsSchedule',
        path: 'formData.formData.additionalData'
      });
    } else if (formDataPayload?.formData?.additionalData?.["Fixed Assets Schedule"]) {
      fixedAssetsSchedule = formDataPayload.formData.additionalData["Fixed Assets Schedule"];
      logger.debug('Fixed assets schedule found in nested path', {
        operation: 'extractFixedAssetsSchedule',
        path: 'formData.additionalData'
      });
    } else if (formDataPayload?.formData?.formData?.["Fixed Assets Schedule"]) {
      fixedAssetsSchedule = formDataPayload.formData.formData["Fixed Assets Schedule"];
      logger.debug('Fixed assets schedule found in formData.formData', {
        operation: 'extractFixedAssetsSchedule',
        path: 'formData.formData'
      });
    } else if (formDataPayload?.formData?.["Fixed Assets Schedule"]) {
      fixedAssetsSchedule = formDataPayload.formData["Fixed Assets Schedule"];
      logger.debug('Fixed assets schedule found in formData', {
        operation: 'extractFixedAssetsSchedule',
        path: 'formData'
      });
    } else if (formDataPayload?.additionalData?.["Fixed Assets Schedule"]) {
      fixedAssetsSchedule = formDataPayload.additionalData["Fixed Assets Schedule"];
      logger.debug('Fixed assets schedule found in additionalData', {
        operation: 'extractFixedAssetsSchedule',
        path: 'additionalData'
      });
    }
    
    if (!fixedAssetsSchedule) {
      logger.warn('Fixed assets schedule not found in payload', {
        operation: 'extractFixedAssetsSchedule',
        templateId,
        payloadKeys: Object.keys(formDataPayload || {})
      });
      return [];
    }
    
    const updates = [];
    
    // Get Fixed Assets mapping from template mapping service
    let categoryRowMapping = templateMappingService.getFixedAssetsMapping(templateId);
    
    if (!categoryRowMapping || Object.keys(categoryRowMapping).length === 0) {
      logger.warn('No fixed assets mapping found for template, using defaults', {
        operation: 'extractFixedAssetsSchedule',
        templateId
      });
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
    
    logger.debug('Using fixed assets mapping', {
      operation: 'extractFixedAssetsSchedule',
      templateId,
      mappingKeys: Object.keys(categoryRowMapping)
    });
    
    // Map frontend category names to backend mapping keys
    const categoryNameMap = {
      "Plant and Machinery": "plant_machinery",
      "Service Equipment": "service_equipment",
      "Civil works & Shed Construction": "shed_civil",
      "Shed Construction and Civil works": "shed_civil",
      "Land": "land",
      "Electrical Items & fittings": "electrical",
      "Electronic Items": "electronic",
      "Furniture and Fittings": "furniture",
      "Vehicles": "vehicles",
      "Live stock": (normalizedTemplateId === 'TERM_LOAN_CC') ? "live_stock" : "other_assets",
      "Other Assets": (normalizedTemplateId === 'TERM_LOAN_CC') ? "other_assets" : "capital_wip",
      "Other Assets (Nil Depreciation)": "other_assets_nil",
      "Non Current Assets (Deposits , Advances etc)": "non_current_assets"
    };
    
    // Determine sheet name based on template type
    const sheetName = (normalizedTemplateId === 'TERM_LOAN_SERVICE_WITHOUT_STOCK' || normalizedTemplateId === 'TERM_LOAN_MANUFACTURING_SERVICE_WITH_STOCK' || normalizedTemplateId === 'TERM_LOAN_CC') 
      ? 'Assumptions' 
      : 'Assumptions.1';

    for (const [categoryName, categoryData] of Object.entries(fixedAssetsSchedule)) {
      if (!categoryData.items || !Array.isArray(categoryData.items)) continue;
      
      // Map frontend category name to backend key
      const mappingKey = categoryNameMap[categoryName] || categoryName.toLowerCase().replace(/\s+/g, '_');
      const startRow = categoryRowMapping[mappingKey];
      
      if (!startRow) {
        logger.warn('Unknown fixed assets category, skipping', {
          operation: 'extractFixedAssetsSchedule',
          categoryName,
          mappingKey
        });
        continue;
      }
      
      // Write each item in the category
      categoryData.items.forEach((item, index) => {
        const row = startRow + index;
        updates.push({ sheet: sheetName, cell: `d${row}`, value: item.description || '' });
        updates.push({ sheet: sheetName, cell: `e${row}`, value: item.amount || 0 });
        logger.debug('Fixed asset item mapped', {
          operation: 'extractFixedAssetsSchedule',
          categoryName,
          row,
          description: item.description,
          amount: item.amount
        });
      });
    }
    
    logger.info('Fixed assets extraction completed', {
      operation: 'extractFixedAssetsSchedule',
      templateId,
      extractedItems: updates.length
    });
    return updates;
  }

  /**
   * Get updates for analysis sheets (Term Loan flow)
   * @param {Object} analysisOptions Analysis options from frontend
   * @param {string} templateId Template ID for sheet name resolution
   * @returns {Array} Array of update objects
   */
  getAnalysisUpdates(analysisOptions, templateId = null) {
    const updates = [];
    if (!analysisOptions || !analysisOptions.extraData) return updates;

    const { sensitivity, bep } = analysisOptions.extraData;
    
    const normalizedTemplateId = templateId ? templateMappingService.normalizeTemplateId(templateId) : null;
    const isTermLoan = normalizedTemplateId === 'TERM_LOAN_SERVICE_WITHOUT_STOCK' || 
                       normalizedTemplateId === 'TERM_LOAN_MANUFACTURING_SERVICE_WITH_STOCK' || 
                       normalizedTemplateId === 'TERM_LOAN_CC';

    const templateConfig = this.getTemplateSheetConfig(templateId);
    const aliasMap = templateConfig?.aliasMap || {};
    
    // Determine target sheet
    // For Term Loans, sensitivity and BEP inputs are in "Assumptions"
    // For others, they are in "Final workings"
    const targetSheet = isTermLoan ? 'Assumptions' : this.resolveSheetAlias('finalworkings', aliasMap);

    // Sensitivity Analysis
    if (sensitivity) {
      if (isTermLoan) {
        // Term Loan: Selling Price Decrease -> d45, Direct Expenses Increase -> d46
        if (sensitivity.sellingPriceDecrease !== undefined) {
          updates.push({
            sheet: targetSheet,
            cell: 'd45',
            value: sensitivity.sellingPriceDecrease
          });
        }
        if (sensitivity.directExpensesIncrease !== undefined) {
          updates.push({
            sheet: targetSheet,
            cell: 'd46',
            value: sensitivity.directExpensesIncrease
          });
        }
      } else {
        // Legacy/Other: Selling Price Decrease -> Cell H54
        if (sensitivity.sellingPriceDecrease !== undefined) {
          updates.push({
            sheet: targetSheet,
            cell: 'H54',
            value: sensitivity.sellingPriceDecrease
          });
        }
      }
    }

    // BEP: Product Manufactured (Units) -> Cell E63, Selling Price per Unit -> Cell E64, 
    // Selling Price Growth -> Cell D65, Plant Operating Capacity per Month -> Cell E66
    if (bep) {
      // For Term Loans, these are also in Assumptions (if we decide to move them)
      // For now, keeping them as they were but using targetSheet
      if (bep.productManufactured) {
        updates.push({ sheet: targetSheet, cell: 'E63', value: bep.productManufactured });
      }
      if (bep.sellingPricePerUnit) {
        updates.push({ sheet: targetSheet, cell: 'E64', value: bep.sellingPricePerUnit });
      }
      if (bep.sellingPriceGrowth) {
        updates.push({ sheet: targetSheet, cell: 'D65', value: bep.sellingPriceGrowth });
      }
      if (bep.plantCapacity) {
        updates.push({ sheet: targetSheet, cell: 'E66', value: bep.plantCapacity });
      }
    }

    return updates;
  }

  /**
   * Extract header fields (Proprietor, Sector, Nature of Business) from payload
   * based on template-specific cell mappings.
   */
  extractHeaderFields(cellData, templateId) {
    const normalizedTemplateId = templateMappingService.normalizeTemplateId(templateId);
    
    // Mapping of header fields to cell IDs for each template
    const headerMapping = {
      'CC1': { proprietor: 'i6', sector: 'i8', natureOfBusiness: 'i9' },
      'CC2': { proprietor: 'i5', sector: 'i7', natureOfBusiness: 'i8' },
      'CC3': { proprietor: 'i6', sector: 'i8', natureOfBusiness: 'i9' },
      'CC4': { proprietor: 'i4', sector: 'i5', natureOfBusiness: 'i6' },
      'CC5': { proprietor: 'i5', sector: 'i7', natureOfBusiness: 'i8' },
      'CC6': { proprietor: 'i5', sector: 'i7', natureOfBusiness: 'i8' },
      'TERM_LOAN_CC': { proprietor: 'i8', sector: 'i14', natureOfBusiness: 'i15' },
      'TERM_LOAN_SERVICE_WITHOUT_STOCK': { proprietor: 'i8', sector: 'i14', natureOfBusiness: 'i15' },
      'TERM_LOAN_MANUFACTURING_SERVICE_WITH_STOCK': { proprietor: 'i8', sector: 'i14', natureOfBusiness: 'i15' }
    };

    const mapping = headerMapping[normalizedTemplateId];
    if (!mapping) {
      return {};
    }

    const result = {};
    if (cellData[mapping.proprietor]) {
      result.proprietor = cellData[mapping.proprietor];
    }
    if (cellData[mapping.sector]) {
      result.sector = cellData[mapping.sector];
    }
    if (cellData[mapping.natureOfBusiness]) {
      result.natureOfBusiness = cellData[mapping.natureOfBusiness];
    }

    return result;
  }

  // ────────────────────────────────────────────────────────────────
  //  Main entry: apply data and extract JSON by calling Python script
  async applyFormDataAndCalculate(templateId, formDataPayload) {
    try {
      logger.info('Starting Python calculation', {
        operation: 'applyFormDataAndCalculate',
        templateId
      });

      // Extract cell data from the payload (with template-specific filtering)
      const cellData = this.extractFormData(formDataPayload, templateId);
      logger.debug('Cell data extracted', {
        operation: 'applyFormDataAndCalculate',
        templateId,
        cellCount: Object.keys(cellData).length
      });

      const updates = [];
      // Normalize template ID for sheet name determination
      const normalizedTemplateId = templateMappingService.normalizeTemplateId(templateId);
      
      // Determine sheet name based on template type
      const sheetName = (normalizedTemplateId === 'TERM_LOAN_SERVICE_WITHOUT_STOCK' || normalizedTemplateId === 'TERM_LOAN_MANUFACTURING_SERVICE_WITH_STOCK' || normalizedTemplateId === 'TERM_LOAN_CC') 
        ? 'Assumptions' 
        : 'Assumptions.1';
      
      for (const [cell, value] of Object.entries(cellData)) {
        updates.push({ sheet: sheetName, cell, value });
      }
      logger.debug('Cell updates prepared', {
        operation: 'applyFormDataAndCalculate',
        templateId,
        updateCount: updates.length
      });
      
      // Extract and add Fixed Assets Schedule items (pass templateId for correct mapping)
      const fixedAssetsUpdates = this.extractFixedAssetsSchedule(formDataPayload, templateId);
      updates.push(...fixedAssetsUpdates);
      logger.debug('Fixed assets updates added', {
        operation: 'applyFormDataAndCalculate',
        templateId,
        fixedAssetsUpdateCount: fixedAssetsUpdates.length
      });
      
      // Extract header fields for the report
      const headerFields = this.extractHeaderFields(cellData, templateId);
      logger.debug('Header fields extracted', {
        operation: 'applyFormDataAndCalculate',
        templateId,
        ...headerFields
      });
      
      const inputData = {
        updates,
        recalculate: false, // Let Excel handle automatic calculation
        skipPdf: true, // Optimization: Skip PDF generation for form application
        ...headerFields
      };
      logger.debug('Input data prepared for Python script', {
        operation: 'applyFormDataAndCalculate',
        templateId,
        totalUpdates: updates.length
      });

      // Resolve template path (handle different naming conventions)
      const templatePath = this.resolveTemplatePath(templateId);
      logger.debug('Template path resolved', {
        operation: 'applyFormDataAndCalculate',
        templateId,
        templatePath
      });
      
      const scriptPath = path.join(this.pythonEnginePath, 'excel_calculator.py');
      logger.debug('Python script path configured', {
        operation: 'applyFormDataAndCalculate',
        scriptPath
      });
      
      logger.debug('Executing Python script', {
        operation: 'applyFormDataAndCalculate',
        templateId,
        pythonExecutable: this.pythonExecutable
      });
      const result = await this.runPythonScript(scriptPath, [templatePath, JSON.stringify(inputData)]);
      
      logger.debug('Python script execution completed', {
        operation: 'applyFormDataAndCalculate',
        templateId
      });
      const excelResult = this.transformPythonResult(JSON.parse(result));

      // PDF is now generated directly in Python, no need for separate generation
      logger.info('PDF generated directly from Excel sheet', {
        operation: 'applyFormDataAndCalculate',
        templateId
      });

      return excelResult;

    } catch (error) {
      logger.error('Error during Python script execution', {
        operation: 'applyFormDataAndCalculate',
        templateId,
        error: error.message,
        stack: error.stack
      });
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
      'Format CC6': 'format CC6.xlsx',
      'TERM_LOAN_SERVICE_WITHOUT_STOCK': 'Term loan (Service sector without stock).xls',
      'TERM_LOAN_MANUFACTURING_SERVICE_WITH_STOCK': 'Term Loan (Manufacturing & Service Sector with stock).xls',
      'TERM_LOAN_CC': 'Term Loan + CC Loan final.xls'
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
  async applyUpdatesAndCalculate(templateId, updatesPayload = {}) {
    let tempWorkbookPath = null;
    try {
      logger.info('Starting Python calculation (applyUpdates)', {
        operation: 'applyUpdatesAndCalculate',
        templateId
      });

      const {
        updates = [],
        recalculate = false,
        baseExcelPath = null,
        existingExcelBuffer = null
      } = updatesPayload || {};

      // Expecting updatesPayload = { updates: [{sheet, cell, value}, ...], recalculate?: boolean }
      const inputData = {
        updates: Array.isArray(updates) ? updates : [],
        recalculate: Boolean(recalculate ?? false), // Default to false, let Excel auto-calculate
        skipPdf: true, // Optimization: Skip PDF generation for updates
      };

      let workbookPath = baseExcelPath;
      if (!workbookPath && existingExcelBuffer) {
        const bufferData = Buffer.isBuffer(existingExcelBuffer)
          ? existingExcelBuffer
          : Buffer.from(existingExcelBuffer, 'base64');
        workbookPath = await this.saveBufferToTempExcel(bufferData, templateId);
        tempWorkbookPath = workbookPath;
        logger.debug('Using staged Excel workbook for updates', {
          operation: 'applyUpdatesAndCalculate',
          templateId,
          workbookPath
        });
      }

      if (!workbookPath) {
        workbookPath = this.resolveTemplatePath(templateId);
      }

      const scriptPath = path.join(this.pythonEnginePath, 'excel_calculator.py');

      const result = await this.runPythonScript(scriptPath, [workbookPath, JSON.stringify(inputData)]);
      logger.debug('Python script execution completed (applyUpdates)', {
        operation: 'applyUpdatesAndCalculate',
        templateId
      });
      return this.transformPythonResult(JSON.parse(result));
    } catch (error) {
      logger.error('Error during Python script execution (applyUpdates)', {
        operation: 'applyUpdatesAndCalculate',
        templateId,
        error: error.message,
        stack: error.stack
      });
      throw new Error('Failed to calculate Excel data using Python engine.');
    } finally {
      if (tempWorkbookPath) {
        try {
          await fsPromises.unlink(tempWorkbookPath);
        } catch (cleanupError) {
          logger.warn('Failed to remove temporary staged Excel file', {
            operation: 'applyUpdatesAndCalculate',
            templateId,
            error: cleanupError.message
          });
        }
      }
    }
  }

  // Generate full AI-enhanced report (Grok only, simplified)
  async generateFullReport(templateId, formDataPayload, apiKey, options = {}) {
    try {
      logger.info('Starting full report generation with Grok', {
        operation: 'generateFullReport',
        templateId
      });

      // Fetch template configuration from database for hidden sheets
      const dbConfig = await TemplateConfig.findOne({ template_id: templateId });
      const excludedSheets = dbConfig?.after_generate_hide || [];
      const defaultFullReportSheets = dbConfig?.full_report_sheets || [];

      // Determine which sheets to include in the report
      let sheetsToInclude = options?.selectedSheets || options?.sheets;
      
      // If no sheets specified, use defaults from DB
      if (!sheetsToInclude && defaultFullReportSheets.length > 0) {
        sheetsToInclude = [...defaultFullReportSheets];
      }

      // If analysis options are provided, merge those sheets as well
      if (options?.analysisOptions?.selectedSheets) {
        const analysisSheets = options.analysisOptions.selectedSheets;
        if (sheetsToInclude) {
          // Merge and remove duplicates
          sheetsToInclude = [...new Set([...sheetsToInclude, ...analysisSheets])];
        } else {
          sheetsToInclude = analysisSheets;
        }
      }

      const selectedSheets = this.normalizeSelectedSheets(sheetsToInclude, templateId);

      // Always use Grok API key
      let finalApiKey = apiKey || process.env.GROK_API_KEY || process.env.XAI_API_KEY;
      if (!finalApiKey) {
        throw new Error('Grok API key is required. Set GROK_API_KEY or XAI_API_KEY environment variable or provide in request.');
      }

      // Extract cell data from the payload (with template-specific filtering)
      const cellData = this.extractFormData(formDataPayload, templateId);

      // Convert cell data to updates array
      const updates = [];
      // Normalize template ID for sheet name determination
      const normalizedTemplateId = templateMappingService.normalizeTemplateId(templateId);
      
      // Determine sheet name based on template type
      const sheetName = (normalizedTemplateId === 'TERM_LOAN_SERVICE_WITHOUT_STOCK' || normalizedTemplateId === 'TERM_LOAN_MANUFACTURING_SERVICE_WITH_STOCK' || normalizedTemplateId === 'TERM_LOAN_CC') 
        ? 'Assumptions' 
        : 'Assumptions.1';

      for (const [cell, value] of Object.entries(cellData)) {
        updates.push({ sheet: sheetName, cell, value });
      }

      // Extract and add Fixed Assets Schedule items (pass templateId for correct mapping)
      const fixedAssetsUpdates = this.extractFixedAssetsSchedule(formDataPayload, templateId);
      updates.push(...fixedAssetsUpdates);

      // Add Analysis Options updates (Term Loan flow)
      if (options?.analysisOptions) {
        const analysisUpdates = this.getAnalysisUpdates(options.analysisOptions, templateId);
        updates.push(...analysisUpdates);
      }

      logger.debug('Excel updates prepared for AI report generation', {
        operation: 'generateFullReport',
        templateId,
        totalUpdates: updates.length,
        cellUpdates: updates.length - fixedAssetsUpdates.length,
        fixedAssetsUpdates: fixedAssetsUpdates.length,
        selectedSheets
      });

      // Build input data for Python script (simplified, Grok only)
      const inputData = {
        updates,
        recalculate: false,
        generateFullReport: true,
        grokApiKey: finalApiKey,  // Always use Grok
        skipHtmlGeneration: true,   // Skip unnecessary HTML generation
        signaturePath: options?.signaturePath || null,
        excludedSheets: excludedSheets // Pass dynamic excluded sheets from DB
      };
      if (selectedSheets) {
        inputData.selectedSheets = selectedSheets;
      }

      const templatePath = this.resolveTemplatePath(templateId);
      const scriptPath = path.join(this.pythonEnginePath, 'excel_calculator.py');

      const result = await this.runPythonScript(scriptPath, [templatePath, JSON.stringify(inputData)]);

      logger.info('Full report generation completed', {
        operation: 'generateFullReport',
        templateId
      });
      logger.debug('Python script output preview', {
        operation: 'generateFullReport',
        templateId,
        outputPreview: result.substring(0, 500)
      });

      let parsedResult;
      try {
        parsedResult = JSON.parse(result);
      } catch (parseError) {
        logger.error('Failed to parse Python output as JSON', {
          operation: 'generateFullReport',
          templateId,
          parseError: parseError.message,
          rawOutputLength: result.length
        });
        throw new Error(`Failed to parse Python script output: ${parseError.message}`);
      }

      const excelResult = this.transformPythonResult(parsedResult);

      return excelResult;

    } catch (error) {
      logger.error('Error during full report generation', {
        operation: 'generateFullReport',
        templateId,
        error: error.message,
        stack: error.stack
      });
      throw error;
    }
  }

  // Generate full AI-enhanced report from existing Excel file
  async generateFullReportFromFile(excelFilePath, templateId, formDataPayload, apiKey, aiProvider = 'grok', options = {}) {
    try {
      logger.info('Starting full report generation from file', {
        operation: 'generateFullReportFromFile',
        templateId,
        excelFilePath,
        aiProvider
      });

      // Fetch template configuration from database for hidden sheets
      const dbConfig = await TemplateConfig.findOne({ template_id: templateId });
      const excludedSheets = dbConfig?.after_generate_hide || [];
      const defaultFullReportSheets = dbConfig?.full_report_sheets || [];

      // Determine which sheets to include in the report
      let sheetsToInclude = options?.selectedSheets || options?.sheets;
      
      // If no sheets specified, use defaults from DB
      if (!sheetsToInclude && defaultFullReportSheets.length > 0) {
        sheetsToInclude = [...defaultFullReportSheets];
      }

      // If analysis options are provided, merge those sheets as well
      if (options?.analysisOptions?.selectedSheets) {
        const analysisSheets = options.analysisOptions.selectedSheets;
        if (sheetsToInclude) {
          // Merge and remove duplicates
          sheetsToInclude = [...new Set([...sheetsToInclude, ...analysisSheets])];
        } else {
          sheetsToInclude = analysisSheets;
        }
      }

      const selectedSheets = this.normalizeSelectedSheets(sheetsToInclude, templateId);

      // Get API key based on provider
      let finalApiKey = apiKey;
      if (!finalApiKey) {
        if (aiProvider === 'grok') {
          finalApiKey = process.env.GROK_API_KEY || process.env.XAI_API_KEY;
        } else if (aiProvider === 'gemini') {
          finalApiKey = process.env.GEMINI_API_KEY;
        } else {
          finalApiKey = process.env.PERPLEXITY_API_KEY;
        }
      }

      if (!finalApiKey) {
        throw new Error(`${aiProvider.toUpperCase()} API key is required. Set ${aiProvider.toUpperCase()}_API_KEY environment variable or provide in request.`);
      }

      // For existing file, we might still need to apply analysis options updates
      const updates = [];
      if (options?.analysisOptions) {
        const analysisUpdates = this.getAnalysisUpdates(options.analysisOptions, templateId);
        updates.push(...analysisUpdates);
      }

      logger.debug('Excel updates prepared for AI report generation from file', {
        operation: 'generateFullReportFromFile',
        templateId,
        excelFilePath,
        totalUpdates: updates.length
      });

      // Build input data for Python script
      const inputData = {
        updates,  // Empty updates since Excel is already updated
        recalculate: false,  // Let Excel auto-calculate
        generateFullReport: true,  // Enable full report generation
        skipHtmlGeneration: true,  // Skip HTML generation for full reports (not needed)
        skipJsonExtraction: true,  // Skip JSON data extraction for full reports (not needed)
        signaturePath: options?.signaturePath || null,
        excludedSheets: excludedSheets // Pass dynamic excluded sheets from DB
      };
      if (selectedSheets) {
        inputData.selectedSheets = selectedSheets;
      }

      // Add API key based on provider
      if (aiProvider === 'grok') {
        inputData.grokApiKey = finalApiKey;
      } else if (aiProvider === 'gemini') {
        inputData.geminiApiKey = finalApiKey;
      } else {
        inputData.perplexityApiKey = finalApiKey;
      }

      const scriptPath = path.join(this.pythonEnginePath, 'excel_calculator.py');
      
      const result = await this.runPythonScript(scriptPath, [excelFilePath, JSON.stringify(inputData)]);
      
      logger.info('Full report generation from file completed', {
        operation: 'generateFullReportFromFile',
        templateId,
        excelFilePath
      });
      logger.debug('Python script output preview', {
        operation: 'generateFullReportFromFile',
        templateId,
        outputPreview: result.substring(0, 500)
      });
      
      let parsedResult;
      try {
        parsedResult = JSON.parse(result);
      } catch (parseError) {
        logger.error('Failed to parse Python output as JSON', {
          operation: 'generateFullReportFromFile',
          templateId,
          excelFilePath,
          parseError: parseError.message,
          rawOutputLength: result.length
        });
        throw new Error(`Failed to parse Python script output: ${parseError.message}`);
      }
      
      const excelResult = this.transformPythonResult(parsedResult);

      return excelResult;

    } catch (error) {
      logger.error('Error during full report generation from file', {
        operation: 'generateFullReportFromFile',
        templateId,
        excelFilePath,
        error: error.message,
        stack: error.stack
      });
      throw error; // Rethrow the actual error instead of generic message
    }
  }

  transformPythonResult(rawResult) {
    logger.debug('Transforming Python result', {
      operation: 'transformPythonResult',
      rawResultType: typeof rawResult,
      rawResultKeys: rawResult ? Object.keys(rawResult) : 'null'
    });
    
    try {
      // Try to parse if it's a string
      let result = rawResult;
      if (typeof rawResult === 'string') {
        logger.debug('Parsing JSON string', {
          operation: 'transformPythonResult'
        });
        result = JSON.parse(rawResult);
        logger.debug('JSON parsed successfully', {
          operation: 'transformPythonResult'
        });
      }
      
      logger.debug('Python result fields', {
        operation: 'transformPythonResult',
        success: result?.success,
        error: result?.error
      });
      
      if (!result || result.success === false) {
        const errorMessage = result?.error || 'Python engine returned an error';
        logger.error('Python script failed', {
          operation: 'transformPythonResult',
          errorMessage
        });
        throw new Error(errorMessage);
      }

      const meta = result._meta || {};
      const verificationCopy = meta.verificationCopy ? path.normalize(meta.verificationCopy) : null;
      const verificationFileName = verificationCopy ? path.basename(verificationCopy) : null;
      // Note: Files now stored in Cloudflare R2 cloud storage, not temp folder

      return {
        relativePath: null, // R2 URLs stored in database
        fileName: verificationFileName,
        excelData: result.excelData,
        jsonData: result.jsonData,
        allSheetsData: result.allSheetsData || {},
        formattedWCData: result.formattedWCData || {},
        htmlContent: result.htmlContent,
        htmlJsonData: result.htmlJsonData || {},
        pdfData: result.pdfData,
        pdfFileName: result.pdfFileName,
        fullReportData: result.fullReportData,
        fullReportFileName: result.fullReportFileName,
        meta: result._meta || {}
      };
      
    } catch (parseError) {
      const rawResultPreview = typeof rawResult === 'string' 
        ? rawResult.substring(0, 500) 
        : (rawResult ? JSON.stringify(rawResult).substring(0, 500) : 'null');

      logger.error('JSON parse error in transformPythonResult', {
        operation: 'transformPythonResult',
        parseError: parseError.message,
        rawResultPreview
      });
      
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
      logger.debug('Executing Python script', {
        operation: 'runPythonScript',
        pythonExecutable: this.pythonExecutable,
        scriptPath,
        args: args.join(' ')
      });
      const env = { ...process.env, TEMP_DIR: this.tempDir };
      const pythonProcess = spawn(this.pythonExecutable, [scriptPath, ...args], { env });

      let stdout = '';
      let stderr = '';

      pythonProcess.stdout.on('data', (data) => {
        stdout += data.toString();
      });

      pythonProcess.stderr.on('data', (data) => {
        const stderrText = data.toString();
        stderr += stderrText;
        // Log Python stderr in real-time for debugging
        logger.debug('Python stderr output', {
          operation: 'runPythonScript',
          stderrText: stderrText.trim()
        });
      });

      pythonProcess.on('close', (code) => {
        if (code !== 0) {
          logger.error('Python script exited with error', {
            operation: 'runPythonScript',
            exitCode: code,
            stderr: stderr.trim()
          });
          return reject(new Error(`Python script failed with code ${code}: ${stderr}`));
        }
        resolve(stdout);
      });

      pythonProcess.on('error', (err) => {
        logger.error('Failed to start Python process', {
          operation: 'runPythonScript',
          error: err.message,
          stack: err.stack
        });
        reject(err);
      });
    });
  }

  // Housekeeping: delete any remaining temp files (legacy - files now stored in R2 cloud)
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

  async saveBufferToTempExcel(buffer, templateId) {
    const safeTemplateId = (templateId || 'template').replace(/[^a-z0-9-_]/gi, '_');
    const tempFileName = `${safeTemplateId}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}.xlsx`;
    const tempFilePath = path.join(this.tempDir, tempFileName);
    await fsPromises.mkdir(this.tempDir, { recursive: true });
    await fsPromises.writeFile(tempFilePath, buffer);
    return tempFilePath;
  }

  normalizeSelectedSheets(sheetList, templateId = null) {
    if (!Array.isArray(sheetList)) {
      return null;
    }

    const templateConfig = this.getTemplateSheetConfig(templateId);
    const aliasMap = templateConfig?.aliasMap || {};
    const normalizedMap = new Map();

    sheetList.forEach((rawSheet) => {
      if (typeof rawSheet !== 'string') {
        return;
      }
      const trimmed = rawSheet.trim();
      if (!trimmed.length) {
        return;
      }
      const canonicalName = this.resolveSheetAlias(trimmed, aliasMap);
      const sheetKey = this.normalizeSheetIdentifier(canonicalName);
      if (!sheetKey || normalizedMap.has(sheetKey)) {
        return;
      }
      normalizedMap.set(sheetKey, canonicalName);
    });

    return normalizedMap.size ? Array.from(normalizedMap.values()) : null;
  }

  getTemplateSheetConfig(templateId) {
    if (!templateId) {
      return null;
    }
    const normalizedTemplateId = templateMappingService.normalizeTemplateId(templateId);
    const lookupKey = (normalizedTemplateId || templateId || '').toUpperCase();
    return TEMPLATE_SHEET_CONFIG[lookupKey] || null;
  }

  normalizeSheetIdentifier(name) {
    if (!name || typeof name !== 'string') {
      return '';
    }
    return name.replace(/[^a-z0-9]/gi, '').toLowerCase();
  }

  resolveSheetAlias(sheetName, aliasMap = {}) {
    const key = this.normalizeSheetIdentifier(sheetName);
    if (key && aliasMap[key]) {
      return aliasMap[key];
    }
    return sheetName;
  }
}

module.exports = new ExcelCalculationService();
