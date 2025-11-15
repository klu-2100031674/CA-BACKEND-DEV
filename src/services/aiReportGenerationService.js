const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs').promises;

/**
 * AI Report Generation Service
 * Orchestrates the complete report generation pipeline:
 * 1. Generate PDFs from all Excel sheets
 * 2. Extract computed data from Excel
 * 3. Generate AI content using Gemini API
 * 4. Merge everything into final report
 */
class AIReportGenerationService {
  constructor() {
    this.pythonEnginePath = path.join(__dirname, '../python-engine');
    this.pythonExecutable = 'C:\\Users\\jithe\\AppData\\Local\\Programs\\Python\\Python312\\python.exe';
    this.tempDir = path.join(__dirname, '../../temp');
    this.geminiApiKey = process.env.GEMINI_API_KEY || null;
  }

  /**
   * Set Gemini API key
   */
  setGeminiApiKey(apiKey) {
    this.geminiApiKey = apiKey;
  }

  /**
   * Generate full AI-enhanced report
   * @param {string} excelFilePath - Path to the calculated Excel file
   * @param {string} templateId - Template identifier
   * @param {object} excelData - Computed data from Excel
   * @returns {Promise<object>} - Report generation result
   */
  async generateFullReport(excelFilePath, templateId, excelData = {}) {
    try {
      console.log('[AIReportService] Starting full report generation');
      console.log(`  Excel file: ${excelFilePath}`);
      console.log(`  Template ID: ${templateId}`);

      // Validate Gemini API key
      if (!this.geminiApiKey) {
        throw new Error('Gemini API key not configured. Set GEMINI_API_KEY environment variable.');
      }

      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const pdfsDir = path.join(this.tempDir, `${templateId}_pdfs_${timestamp}`);
      const finalReportPath = path.join(this.tempDir, `${templateId}_full_report_${timestamp}.pdf`);

      // Ensure directories exist
      await fs.mkdir(pdfsDir, { recursive: true });

      // Step 1: Generate PDFs for all Excel sheets
      console.log('[AIReportService] Step 1: Generating PDFs from Excel sheets...');
      const pdfGeneration = await this.generateExcelSheetPDFs(excelFilePath, pdfsDir);

      if (!pdfGeneration.success) {
        throw new Error('Failed to generate Excel sheet PDFs');
      }

      console.log(`[AIReportService] Generated ${pdfGeneration.success_count} PDFs from Excel sheets`);

      // Step 2: Save Excel data to JSON for Python script
      const excelDataPath = path.join(this.tempDir, `${templateId}_data_${timestamp}.json`);
      await fs.writeFile(excelDataPath, JSON.stringify(excelData, null, 2));

      // Step 3: Generate AI-enhanced report using Python
      console.log('[AIReportService] Step 2: Generating AI-enhanced content and merging PDFs...');
      const reportGeneration = await this.generateAIReport(pdfsDir, excelDataPath, finalReportPath);

      if (!reportGeneration.success) {
        throw new Error('Failed to generate AI-enhanced report');
      }

      // Step 4: Read final PDF and encode to base64
      const pdfBuffer = await fs.readFile(finalReportPath);
      const pdfBase64 = pdfBuffer.toString('base64');

      // Step 5: Cleanup temporary files
      try {
        await fs.unlink(excelDataPath);
        // Optionally cleanup PDF directory (keep for debugging)
        // await fs.rm(pdfsDir, { recursive: true });
      } catch (cleanupError) {
        console.warn('[AIReportService] Cleanup warning:', cleanupError.message);
      }

      console.log('[AIReportService] ✅ Full report generation complete');

      return {
        success: true,
        message: 'AI-enhanced report generated successfully',
        data: {
          reportPath: finalReportPath,
          reportFileName: path.basename(finalReportPath),
          reportBase64: pdfBase64,
          fileSize: pdfBuffer.length,
          excelPDFsGenerated: pdfGeneration.success_count,
          aiSectionsGenerated: reportGeneration.ai_sections_generated || [],
          url: `/temp/${path.basename(finalReportPath)}`
        }
      };

    } catch (error) {
      console.error('[AIReportService] Error generating full report:', error);
      throw error;
    }
  }

  /**
   * Generate individual PDFs for all Excel sheets
   * @param {string} excelPath - Path to Excel file
   * @param {string} outputDir - Directory to save PDFs
   * @returns {Promise<object>} - Generation result
   */
  async generateExcelSheetPDFs(excelPath, outputDir) {
    const scriptPath = path.join(this.pythonEnginePath, 'excel_calculator.py');

    // Create a Python script call to generate all sheet PDFs
    const pythonCode = `
import sys
import json
from excel_calculator import generate_pdfs_for_all_sheets

excel_path = sys.argv[1]
output_dir = sys.argv[2]

result = generate_pdfs_for_all_sheets(excel_path, output_dir)
print(json.dumps(result))
`;

    const tempScriptPath = path.join(this.tempDir, 'generate_sheet_pdfs.py');
    await fs.writeFile(tempScriptPath, pythonCode);

    try {
      const output = await this.runPythonScript(tempScriptPath, [excelPath, outputDir]);
      const result = JSON.parse(output);

      await fs.unlink(tempScriptPath); // Cleanup temp script
      return result;

    } catch (error) {
      console.error('[AIReportService] Error generating Excel sheet PDFs:', error);
      throw error;
    }
  }

  /**
   * Generate AI-enhanced report using Python pdf_report_generator
   * @param {string} pdfsDir - Directory containing Excel sheet PDFs
   * @param {string} excelDataPath - Path to JSON file with Excel data
   * @param {string} outputPath - Output path for final report
   * @returns {Promise<object>} - Generation result
   */
  async generateAIReport(pdfsDir, excelDataPath, outputPath) {
    const scriptPath = path.join(this.pythonEnginePath, 'pdf_report_generator.py');

    const args = [
      '--api-key', this.geminiApiKey,
      '--excel-pdfs-dir', pdfsDir,
      '--output', outputPath,
      '--excel-data', excelDataPath
    ];

    try {
      const output = await this.runPythonScript(scriptPath, args);
      const result = JSON.parse(output);
      return result;

    } catch (error) {
      console.error('[AIReportService] Error generating AI report:', error);
      throw error;
    }
  }

  /**
   * Run a Python script and return its output
   * @param {string} scriptPath - Path to Python script
   * @param {Array<string>} args - Script arguments
   * @returns {Promise<string>} - Script output
   */
  runPythonScript(scriptPath, args) {
    return new Promise((resolve, reject) => {
      const pythonProcess = spawn(this.pythonExecutable, [scriptPath, ...args]);

      let stdout = '';
      let stderr = '';

      pythonProcess.stdout.on('data', (data) => {
        stdout += data.toString();
      });

      pythonProcess.stderr.on('data', (data) => {
        const stderrText = data.toString();
        stderr += stderrText;
        // Log Python stderr for debugging
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
        console.error('Failed to start Python process:', err);
        reject(err);
      });
    });
  }

  /**
   * Check if Gemini API key is configured
   * @returns {boolean}
   */
  isConfigured() {
    return this.geminiApiKey !== null;
  }

  /**
   * Get service status
   * @returns {object}
   */
  getStatus() {
    return {
      geminiConfigured: this.isConfigured(),
      pythonExecutable: this.pythonExecutable,
      pythonEnginePath: this.pythonEnginePath,
      tempDir: this.tempDir
    };
  }
}

module.exports = new AIReportGenerationService();
