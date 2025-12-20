/**
 * PDF Generation Service
 * 
 * Handles PDF regeneration from Excel files.
 * Used when admin uploads a revised Excel file and needs to regenerate the PDF.
 */

const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const fsPromises = fs.promises;
const logger = require('../utils/logger');

class PdfGenerationService {
  constructor() {
    this.pythonEnginePath = path.join(__dirname, '../python-engine');
    this.scriptPath = path.join(this.pythonEnginePath, 'pdf_regenerator.py');
    this.tempDir = path.join(__dirname, '../../temp');
    
    // Use virtual environment Python (same logic as excelCalculationService)
    const venvDir = process.platform === 'win32' ? 'Scripts' : 'bin';
    const pythonExe = process.platform === 'win32' ? 'python.exe' : 'python';
    const venvPythonPath = path.join(this.pythonEnginePath, '.venv', venvDir, pythonExe);
    
    // Check if virtual environment Python exists; otherwise, use system Python
    this.pythonExecutable = fs.existsSync(venvPythonPath) ? venvPythonPath : 'python';
    
    logger.debug('PdfGenerationService initialized', {
      operation: 'constructor',
      pythonExecutable: this.pythonExecutable,
      scriptPath: this.scriptPath
    });
  }

  /**
   * Regenerate PDF from an Excel buffer
   * Uses AIReportGenerator for proper sheet ordering and AI content generation.
   * 
   * @param {Object} options
   * @param {Buffer} options.excelBuffer - The Excel file buffer
   * @param {string} options.templateId - Template ID for the report
   * @param {string[]} options.selectedSheets - Array of sheet names to include in PDF (optional)
   * @param {string} options.grokApiKey - Grok API key for AI content generation
   * @param {Object} options.jsonData - JSON data from original report for AI context
   * @param {Object} options.htmlData - HTML data from original report
   * @param {string} options.templateName - Template name (CC1, CC2, TL1, etc.)
   * @param {string} options.signatureUrl - URL of the admin signature image
   * @returns {Promise<{success: boolean, pdfBuffer?: Buffer, pdfFileName?: string, error?: string}>}
   */
  async regeneratePdfFromExcel({ 
    excelBuffer, 
    templateId, 
    selectedSheets = null,
    grokApiKey = null,
    jsonData = null,
    htmlData = null,
    templateName = null,
    signatureUrl = null
  }) {
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const tempExcelPath = path.join(this.tempDir, `${templateId}_regen_${timestamp}.xlsx`);
    let tempSignaturePath = null;
    
    logger.info('Starting PDF regeneration from Excel', {
      operation: 'regeneratePdfFromExcel',
      templateId,
      templateName,
      selectedSheets,
      excelSize: excelBuffer.length,
      hasGrokApiKey: !!grokApiKey,
      hasJsonData: !!jsonData,
      hasSignature: !!signatureUrl
    });

    try {
      // Ensure temp directory exists
      await fsPromises.mkdir(this.tempDir, { recursive: true });
      
      // Write Excel buffer to temp file
      await fsPromises.writeFile(tempExcelPath, excelBuffer);
      
      // Download signature if provided
      if (signatureUrl) {
        try {
          const r2Service = require('./cloudflareR2Service');
          const signatureKey = r2Service.extractKeyFromUrl(signatureUrl);
          if (signatureKey) {
            const signatureBuffer = await r2Service.downloadFile(signatureKey);
            const ext = path.extname(signatureUrl) || '.png';
            tempSignaturePath = path.join(this.tempDir, `sig_${timestamp}${ext}`);
            await fsPromises.writeFile(tempSignaturePath, signatureBuffer);
            logger.debug('Signature downloaded to temp path', { tempSignaturePath });
          }
        } catch (sigError) {
          logger.error('Failed to download signature, proceeding without it', { error: sigError.message });
        }
      }

      logger.debug('Excel file written to temp path', {
        operation: 'regeneratePdfFromExcel',
        tempExcelPath
      });

      // Prepare input JSON for Python script
      const inputData = {
        excel_path: tempExcelPath,
        selected_sheets: selectedSheets,
        grok_api_key: grokApiKey,
        json_data: jsonData,
        html_data: htmlData,
        template_name: templateName || templateId,
        signature_path: tempSignaturePath
      };

      // Run Python script with JSON input
      const result = await this.runPythonScript(this.scriptPath, [JSON.stringify(inputData)]);
      
      // Parse result
      let parsedResult;
      try {
        parsedResult = JSON.parse(result);
      } catch (parseError) {
        logger.error('Failed to parse Python script output', {
          operation: 'regeneratePdfFromExcel',
          error: parseError.message,
          rawResult: result.substring(0, 500)
        });
        throw new Error('Failed to parse PDF generation result');
      }

      if (!parsedResult.success) {
        throw new Error(parsedResult.error || 'PDF regeneration failed');
      }

      // Decode base64 PDF data
      const pdfBuffer = Buffer.from(parsedResult.pdf_base64, 'base64');

      logger.info('PDF regeneration completed successfully', {
        operation: 'regeneratePdfFromExcel',
        templateId,
        sheetsProcessed: parsedResult.sheets_processed,
        aiSectionsGenerated: parsedResult.ai_sections_generated || 0,
        pdfSize: pdfBuffer.length
      });

      return {
        success: true,
        pdfBuffer,
        pdfFileName: parsedResult.pdf_filename,
        sheetsProcessed: parsedResult.sheets_processed,
        aiSectionsGenerated: parsedResult.ai_sections_generated || 0
      };

    } catch (error) {
      logger.error('PDF regeneration failed', {
        operation: 'regeneratePdfFromExcel',
        templateId,
        error: error.message,
        stack: error.stack
      });

      return {
        success: false,
        error: error.message
      };

    } finally {
      // Cleanup temp Excel file
      try {
        if (fs.existsSync(tempExcelPath)) {
          await fsPromises.unlink(tempExcelPath);
        }
        
        // Cleanup temp signature file
        if (tempSignaturePath && fs.existsSync(tempSignaturePath)) {
          await fsPromises.unlink(tempSignaturePath);
        }

        logger.debug('Cleaned up temp files', {
          operation: 'regeneratePdfFromExcel',
          tempExcelPath,
          tempSignaturePath
        });
      } catch (cleanupError) {
        logger.warn('Failed to cleanup temp files', {
          operation: 'regeneratePdfFromExcel',
          error: cleanupError.message
        });
      }
    }
  }

  /**
   * Run a Python script and return its output
   * 
   * @param {string} scriptPath - Path to the Python script
   * @param {string[]} args - Arguments to pass to the script
   * @returns {Promise<string>} - Script stdout
   */
  runPythonScript(scriptPath, args) {
    return new Promise((resolve, reject) => {
      logger.debug('Executing Python script for PDF regeneration', {
        operation: 'runPythonScript',
        pythonExecutable: this.pythonExecutable,
        scriptPath,
        args: args.slice(0, 1).join(' ') + (args.length > 1 ? ' [...]' : '')
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
}

// Export singleton instance
module.exports = new PdfGenerationService();
