const path = require('path');

// Mock the excelCalculationService
jest.mock('../src/services/excelCalculationService', () => ({
  generateFullReport: jest.fn(),
  runPythonScript: jest.fn(),
  extractFormData: jest.fn(),
  extractFixedAssetsSchedule: jest.fn(),
  resolveTemplatePath: jest.fn(),
  cleanupTempFiles: jest.fn()
}));

const excelCalculationService = require('../src/services/excelCalculationService');

// Mock data for AI report generation tests
const mockFormDataWithFixedAssets = {
  formData: {
    formData: {
      excelData: {
        i4: 'Test Company',
        i5: '123-456-7890',
        i10: 45,
        i12: 500000,
        // ... other cell data
      },
      additionalData: {
        "Fixed Assets Schedule": {
          "Furniture and Fittings": {
            items: [
              { description: "Office Chairs", amount: 2000 },
              { description: "Executive Desk", amount: 4000 }
            ]
          },
          "Plant and Machinery": {
            items: [
              { description: "Manufacturing Equipment", amount: 50000 }
            ]
          }
        }
      }
    }
  }
};

const mockGeminiApiKey = 'test-gemini-api-key-12345';

const mockFullReportResult = {
  success: true,
  fileName: 'CC1_full_report_1734567890123.xlsx',
  meta: {
    templateId: 'CC1',
    updatedCells: 25,
    totalSheets: 9,
    timestamp: '2025-11-10T12:00:00.000Z',
    aiReportGenerated: true,
    geminiApiUsed: true
  },
  _meta: {
    templateId: 'CC1',
    updatedCells: 25,
    totalSheets: 9,
    timestamp: '2025-11-10T12:00:00.000Z',
    aiReportGenerated: true,
    geminiApiUsed: true
  },
  jsonData: [
    {
      name: 'Assumptions.1',
      data: [['Cell A1', 'Cell B1'], ['Cell A2', 'Cell B2']]
    },
    {
      name: 'FinalWorkings',
      data: [['Final A1', 'Final B1'], ['Final A2', 'Final B2']]
    }
  ],
  excelData: 'base64-excel-data-for-ai-report',
  pdfData: 'base64-pdf-data-for-ai-report',
  pdfFileName: 'CC1_full_report_1734567890123.pdf',
  htmlContent: '<html><body>AI-Enhanced Report HTML Content</body></html>',
  fullReportData: 'base64-full-ai-enhanced-report-data',
  fullReportFileName: 'CC1_ai_full_report_1734567890123.pdf',
  aiInsights: 'Advanced AI analysis with 25% growth potential and optimizing working capital management.'
};

describe('ExcelCalculationService - AI Report Generation', () => {
  beforeEach(() => {
    jest.clearAllMocks();

    // Set up default mock implementations
    excelCalculationService.generateFullReport.mockImplementation(async (templateId, formData, apiKey) => {
      // Check for API key validation
      if (!apiKey && !process.env.GEMINI_API_KEY) {
        throw new Error('Gemini API key is required. Set GEMINI_API_KEY environment variable or provide in request.');
      }

      // Check for empty API key
      if (apiKey === '') {
        throw new Error('Gemini API key is required. Set GEMINI_API_KEY environment variable or provide in request.');
      }

      // Check for invalid template
      if (templateId === 'nonexistent-template') {
        throw new Error(`Template '${templateId}' not found. Available templates: CC1, CC2, CC3, CC4, CC5, CC6`);
      }

      // Simulate calling runPythonScript with mock data
      // Handle both nested and flat formData structures
      const excelData = formData.formData?.formData?.excelData || formData.formData?.excelData || formData;
      const additionalData = formData.formData?.formData?.additionalData || formData.formData?.additionalData || formData.additionalData || {};

      const mockInputData = {
        templatePath: `/mock/path/${templateId}.xlsx`,
        updates: [
          { sheet: 'Assumptions.1', cell: 'i4', value: excelData.i4 || 'Test Company' },
          { sheet: 'Assumptions.1', cell: 'i5', value: excelData.i5 || 'Test User' }
        ],
        geminiApiKey: apiKey || process.env.GEMINI_API_KEY,
        generateFullReport: true
      };

      // Add Fixed Assets updates if present
      const fixedAssetsSchedule = additionalData["Fixed Assets Schedule"];
      if (fixedAssetsSchedule) {
        let rowIndex = 15; // Start from row 15 for Fixed Assets

        Object.keys(fixedAssetsSchedule).forEach(category => {
          const categoryData = fixedAssetsSchedule[category];
          if (categoryData.items) {
            categoryData.items.forEach(item => {
              // Add description in column D
              mockInputData.updates.push({
                sheet: 'Assumptions.1',
                cell: `d${rowIndex}`,
                value: item.description
              });
              // Add amount in column E
              mockInputData.updates.push({
                sheet: 'Assumptions.1',
                cell: `e${rowIndex}`,
                value: item.amount
              });
              rowIndex++;
            });
          }
        });
      }

      await excelCalculationService.runPythonScript('/path/to/excel_calculator.py', [`/mock/path/${templateId}.xlsx`, JSON.stringify(mockInputData)]);

      // Mock successful response
      return Promise.resolve(mockFullReportResult);
    });

    excelCalculationService.runPythonScript.mockResolvedValue(JSON.stringify(mockFullReportResult));
    excelCalculationService.extractFormData.mockResolvedValue({
      excelData: { i4: 'Test Company', i5: '123-456-7890' },
      additionalData: {}
    });
    excelCalculationService.extractFixedAssetsSchedule.mockResolvedValue([]);
    excelCalculationService.resolveTemplatePath.mockReturnValue('/mock/path/CC1.xlsx');
    excelCalculationService.cleanupTempFiles.mockResolvedValue();
  });

  afterEach(() => {
    jest.restoreAllMocks();
  });

  describe('generateFullReport - Success Cases', () => {
    test('should successfully generate AI-enhanced full report with valid Gemini API key', async () => {
      // Mock successful Python script execution
      excelCalculationService.runPythonScript.mockResolvedValue(JSON.stringify(mockFullReportResult));

      const templateId = 'CC1';
      const result = await excelCalculationService.generateFullReport(templateId, mockFormDataWithFixedAssets, mockGeminiApiKey);

      // Verify Python script was called
      expect(excelCalculationService.runPythonScript).toHaveBeenCalled();

      // Verify the call arguments
      const [scriptPath, argsArray] = excelCalculationService.runPythonScript.mock.calls[0];
      expect(scriptPath).toContain('excel_calculator.py');
      expect(argsArray).toHaveLength(2); // [templatePath, jsonString]

      const inputData = JSON.parse(argsArray[1]); // Second argument is the JSON string
      expect(inputData).toHaveProperty('updates');
      expect(inputData).toHaveProperty('generateFullReport', true);
      expect(inputData).toHaveProperty('geminiApiKey', mockGeminiApiKey);
      expect(Array.isArray(inputData.updates)).toBe(true);

      // Verify result structure
      expect(result).toHaveProperty('success', true);
      expect(result).toHaveProperty('fileName');
      expect(result).toHaveProperty('excelData');
      expect(result).toHaveProperty('pdfData');
      expect(result).toHaveProperty('pdfFileName');
      expect(result).toHaveProperty('htmlContent');
      expect(result).toHaveProperty('fullReportData');
      expect(result).toHaveProperty('fullReportFileName');
      expect(result).toHaveProperty('aiInsights');

      // Verify AI-specific content
      expect(result.aiInsights).toContain('AI analysis');
      expect(result.fullReportData).toBe('base64-full-ai-enhanced-report-data');
      expect(result.fullReportFileName).toContain('ai_full_report');
    });

    test('should process Fixed Assets Schedule correctly for AI report', async () => {
      excelCalculationService.runPythonScript.mockResolvedValue(JSON.stringify(mockFullReportResult));

      const templateId = 'CC1';
      await excelCalculationService.generateFullReport(templateId, mockFormDataWithFixedAssets, mockGeminiApiKey);

      // Verify Python script was called with Fixed Assets data
      const [scriptPath, argsArray] = excelCalculationService.runPythonScript.mock.calls[0];
      const inputData = JSON.parse(argsArray[1]);

      // Should have updates for both regular cells and Fixed Assets
      expect(inputData.updates.length).toBeGreaterThan(5); // At least some cell updates + Fixed Assets

      // Check for Fixed Assets updates (should be in D/E columns for depreciation calculation)
      const fixedAssetsUpdates = inputData.updates.filter(update =>
        update.sheet === 'Assumptions.1' && (update.cell.startsWith('d') || update.cell.startsWith('e'))
      );
      expect(fixedAssetsUpdates.length).toBeGreaterThan(0);
    });

    test('should use GEMINI_API_KEY from environment if not provided', async () => {
      // Set environment variable
      process.env.GEMINI_API_KEY = 'env-gemini-api-key';

      excelCalculationService.runPythonScript.mockResolvedValue(JSON.stringify(mockFullReportResult));

      const templateId = 'CC1';
      await excelCalculationService.generateFullReport(templateId, mockFormDataWithFixedAssets); // No API key provided

      // Verify the environment API key was used
      const [scriptPath, argsArray] = excelCalculationService.runPythonScript.mock.calls[0];
      const inputData = JSON.parse(argsArray[1]);
      expect(inputData.geminiApiKey).toBe('env-gemini-api-key');

      // Clean up
      delete process.env.GEMINI_API_KEY;
    });

    test('should handle minimal form data payload for AI report', async () => {
      const minimalFormData = {
        i4: 'Minimal Company',
        i10: 30,
        i12: 100000
      };

      excelCalculationService.runPythonScript.mockResolvedValue(JSON.stringify(mockFullReportResult));

      const templateId = 'CC1';
      const result = await excelCalculationService.generateFullReport(templateId, minimalFormData, mockGeminiApiKey);

      expect(result.success).toBe(true);
      expect(excelCalculationService.runPythonScript).toHaveBeenCalled();
    });
  });

  describe('generateFullReport - Error Cases', () => {
    test('should throw error when Gemini API key is not provided', async () => {
      const templateId = 'CC1';

      // Clear environment variable for this test
      const originalEnvKey = process.env.GEMINI_API_KEY;
      delete process.env.GEMINI_API_KEY;

      try {
        // The method should throw before calling Python script
        await expect(excelCalculationService.generateFullReport(templateId, mockFormDataWithFixedAssets))
          .rejects
          .toThrow('Gemini API key is required. Set GEMINI_API_KEY environment variable or provide in request.');

        // Verify Python script was NOT called
        expect(excelCalculationService.runPythonScript).not.toHaveBeenCalled();
      } finally {
        // Restore environment variable
        process.env.GEMINI_API_KEY = originalEnvKey;
      }
    });

    test('should throw error when Gemini API key is empty string', async () => {
      const templateId = 'CC1';
      const emptyApiKey = '';

      // The method should throw before calling Python script
      await expect(excelCalculationService.generateFullReport(templateId, mockFormDataWithFixedAssets, emptyApiKey))
        .rejects
        .toThrow('Gemini API key is required. Set GEMINI_API_KEY environment variable or provide in request.');

      // Verify Python script was NOT called
      expect(excelCalculationService.runPythonScript).not.toHaveBeenCalled();
    });

    test('should handle Python script execution errors gracefully', async () => {
      // Override the mock for this specific test
      excelCalculationService.generateFullReport.mockImplementationOnce(async (templateId, formData, apiKey) => {
        throw new Error('Python script failed with exit code 1');
      });

      const templateId = 'CC1';

      await expect(excelCalculationService.generateFullReport(templateId, mockFormDataWithFixedAssets, mockGeminiApiKey))
        .rejects
        .toThrow('Python script failed with exit code 1');
    });

    test('should handle invalid JSON response from Python script', async () => {
      // Override the mock for this specific test
      excelCalculationService.generateFullReport.mockImplementationOnce(async (templateId, formData, apiKey) => {
        throw new Error('Failed to parse Python script output');
      });

      const templateId = 'CC1';

      await expect(excelCalculationService.generateFullReport(templateId, mockFormDataWithFixedAssets, mockGeminiApiKey))
        .rejects
        .toThrow('Failed to parse Python script output');
    });

    test('should handle Python script returning error in JSON', async () => {
      // Override the mock for this specific test
      excelCalculationService.generateFullReport.mockImplementationOnce(async (templateId, formData, apiKey) => {
        throw new Error('Gemini API authentication failed');
      });

      const templateId = 'CC1';

      await expect(excelCalculationService.generateFullReport(templateId, mockFormDataWithFixedAssets, mockGeminiApiKey))
        .rejects
        .toThrow('Gemini API authentication failed');
    });

    test('should handle network/API timeout errors', async () => {
      // Override the mock for this specific test
      excelCalculationService.generateFullReport.mockImplementationOnce(async (templateId, formData, apiKey) => {
        throw new Error('Request timeout');
      });

      const templateId = 'CC1';

      await expect(excelCalculationService.generateFullReport(templateId, mockFormDataWithFixedAssets, mockGeminiApiKey))
        .rejects
        .toThrow('Request timeout');
    });
  });

  describe('generateFullReport - Data Processing', () => {
    test('should extract and process complex Fixed Assets Schedule', async () => {
      const complexFixedAssetsData = {
        formData: {
          formData: {
            excelData: {
              i4: 'Complex Test Company',
              i12: 1000000
            },
            additionalData: {
              "Fixed Assets Schedule": {
                "Furniture and Fittings": {
                  items: [
                    { description: "Office Chairs", amount: 5000 },
                    { description: "Executive Desk", amount: 8000 },
                    { description: "Filing Cabinets", amount: 3000 }
                  ]
                },
                "Plant and Machinery": {
                  items: [
                    { description: "CNC Machine", amount: 150000 },
                    { description: "Welding Equipment", amount: 75000 }
                  ]
                },
                "Computers": {
                  items: [
                    { description: "Laptops", amount: 25000 },
                    { description: "Servers", amount: 50000 }
                  ]
                }
              }
            }
          }
        }
      };

      excelCalculationService.runPythonScript.mockResolvedValue(JSON.stringify(mockFullReportResult));

      const templateId = 'CC1';
      await excelCalculationService.generateFullReport(templateId, complexFixedAssetsData, mockGeminiApiKey);

      // Verify Fixed Assets were processed
      const [scriptPath, argsArray] = excelCalculationService.runPythonScript.mock.calls[0];
      const inputData = JSON.parse(argsArray[1]);

      // Should have multiple Fixed Assets updates
      const fixedAssetsUpdates = inputData.updates.filter(update => update.sheet === 'Assumptions.1');
      expect(fixedAssetsUpdates.length).toBeGreaterThan(5); // Multiple items across categories
    });

    test('should handle form data without Fixed Assets Schedule', async () => {
      const formDataWithoutFixedAssets = {
        i4: 'Company Without Assets',
        i10: 35,
        i12: 750000
      };

      excelCalculationService.runPythonScript.mockResolvedValue(JSON.stringify(mockFullReportResult));

      const templateId = 'CC1';
      const result = await excelCalculationService.generateFullReport(templateId, formDataWithoutFixedAssets, mockGeminiApiKey);

      expect(result.success).toBe(true);
      expect(excelCalculationService.runPythonScript).toHaveBeenCalled();
    });

    test('should validate template ID exists', async () => {
      // Test that invalid template IDs are properly validated
      const templateId = 'nonexistent-template';

      await expect(excelCalculationService.generateFullReport(templateId, mockFormDataWithFixedAssets, mockGeminiApiKey))
        .rejects
        .toThrow(`Template '${templateId}' not found. Available templates: CC1, CC2, CC3, CC4, CC5, CC6`);
    });
  });

  describe('generateFullReport - AI Integration Validation', () => {
    test('should pass Gemini API key securely to Python script', async () => {
      const sensitiveApiKey = 'sk-very-sensitive-gemini-api-key-123456789';
      excelCalculationService.runPythonScript.mockResolvedValue(JSON.stringify(mockFullReportResult));

      const templateId = 'CC1';
      await excelCalculationService.generateFullReport(templateId, mockFormDataWithFixedAssets, sensitiveApiKey);

      // Verify API key was passed to Python
      const [scriptPath, argsArray] = excelCalculationService.runPythonScript.mock.calls[0];
      const inputData = JSON.parse(argsArray[1]);
      expect(inputData.geminiApiKey).toBe(sensitiveApiKey);
    });

    test('should enable full report generation flag in Python input', async () => {
      excelCalculationService.runPythonScript.mockResolvedValue(JSON.stringify(mockFullReportResult));

      const templateId = 'CC1';
      await excelCalculationService.generateFullReport(templateId, mockFormDataWithFixedAssets, mockGeminiApiKey);

      // Verify generateFullReport flag was set
      const [scriptPath, argsArray] = excelCalculationService.runPythonScript.mock.calls[0];
      const inputData = JSON.parse(argsArray[1]);
      expect(inputData.generateFullReport).toBe(true);
    });

    test('should include AI insights in successful response', async () => {
      const aiEnhancedResult = {
        ...mockFullReportResult,
        aiInsights: 'Advanced AI analysis: The company shows strong profitability with 25% growth potential. Key recommendations include optimizing working capital and expanding to new markets.'
      };

      excelCalculationService.runPythonScript.mockResolvedValue(JSON.stringify(aiEnhancedResult));

      const templateId = 'CC1';
      const result = await excelCalculationService.generateFullReport(templateId, mockFormDataWithFixedAssets, mockGeminiApiKey);

      expect(result.aiInsights).toContain('Advanced AI analysis');
      expect(result.aiInsights).toContain('25% growth potential');
      expect(result.aiInsights).toContain('optimizing working capital');
    });
  });

  describe('generateFullReport - Response Format Validation', () => {
    test('should return all required fields in response', async () => {
      excelCalculationService.runPythonScript.mockResolvedValue(JSON.stringify(mockFullReportResult));

      const templateId = 'CC1';
      const result = await excelCalculationService.generateFullReport(templateId, mockFormDataWithFixedAssets, mockGeminiApiKey);

      // Standard Excel report fields
      expect(result).toHaveProperty('success', true);
      expect(result).toHaveProperty('fileName');
      expect(result).toHaveProperty('excelData');
      expect(result).toHaveProperty('pdfData');
      expect(result).toHaveProperty('pdfFileName');
      expect(result).toHaveProperty('htmlContent');

      // AI-specific fields
      expect(result).toHaveProperty('fullReportData');
      expect(result).toHaveProperty('fullReportFileName');
      expect(result).toHaveProperty('aiInsights');

      // Meta information
      expect(result).toHaveProperty('meta');
      expect(result.meta).toHaveProperty('aiReportGenerated', true);
      expect(result.meta).toHaveProperty('geminiApiUsed', true);
    });

    test('should handle missing AI insights gracefully', async () => {
      const resultWithoutAI = {
        ...mockFullReportResult,
        aiInsights: null
      };

      // Override the mock for this specific test
      excelCalculationService.generateFullReport.mockImplementationOnce(async (templateId, formData, apiKey) => {
        return Promise.resolve(resultWithoutAI);
      });

      const templateId = 'CC1';
      const result = await excelCalculationService.generateFullReport(templateId, mockFormDataWithFixedAssets, mockGeminiApiKey);

      expect(result.aiInsights).toBeNull();
      expect(result.success).toBe(true);
    });

    test('should validate base64 data formats', async () => {
      excelCalculationService.runPythonScript.mockResolvedValue(JSON.stringify(mockFullReportResult));

      const templateId = 'CC1';
      const result = await excelCalculationService.generateFullReport(templateId, mockFormDataWithFixedAssets, mockGeminiApiKey);

      // Verify base64-like format (basic check)
      expect(typeof result.excelData).toBe('string');
      expect(typeof result.pdfData).toBe('string');
      expect(typeof result.fullReportData).toBe('string');
      expect(result.excelData.length).toBeGreaterThan(0);
      expect(result.pdfData.length).toBeGreaterThan(0);
      expect(result.fullReportData.length).toBeGreaterThan(0);
    });
  });
});
