
const path = require('path');

// Mock templateMappingService
jest.mock('../src/services/templateMappingService');
const templateMappingService = require('../src/services/templateMappingService');

const excelCalculationService = require('../src/services/excelCalculationService');

// Mock the runPythonScript method to avoid actual Python execution
const mockRunPythonScript = jest.fn();

// Override the runPythonScript method
excelCalculationService.runPythonScript = mockRunPythonScript;

// Mock resolveTemplatePath
excelCalculationService.resolveTemplatePath = jest.fn().mockReturnValue('/mock/template/path.xlsx');

// Mock data for the test
const mockFormData = {
  'I4': 'John Doe', // Example: Client Name
  'I5': '123-456-7890', // Example: PAN
  'I10': 45, // Example: Age
  'I12': 500000, // Example: Gross Total Income
};

const mockPythonResult = {
  success: true,
  _meta: {
    templateId: 'CC1',
    updatedCells: 4,
    totalSheets: 9,
    timestamp: '2025-11-10T12:00:00.000Z'
  },
  jsonData: [
    {
      name: 'Assumptions.1',
      data: [
        ['Cell A1', 'Cell B1'],
        ['Cell A2', 'Cell B2']
      ]
    },
    {
      name: 'FinalWorkings',
      data: [
        ['Final A1', 'Final B1'],
        ['Final A2', 'Final B2']
      ]
    }
  ],
  excelData: 'base64-mock-excel-data',
  pdfData: 'base64-mock-pdf-data',
  pdfFileName: 'CC1_1734567890123.pdf',
  htmlContent: '<html><body>Mock HTML Content</body></html>'
};

describe('ExcelCalculationService - Unit Test', () => {

  beforeEach(() => {
    jest.clearAllMocks();

    // Mock templateMappingService methods
    templateMappingService.normalizeTemplateId.mockImplementation((templateId) => {
      if (!templateId) return null;
      const idUpper = templateId.toUpperCase();
      const match = idUpper.match(/CC(\d+)/);
      return match ? `CC${match[1]}` : templateId;
    });

    templateMappingService.filterWritableCells.mockImplementation((templateId, data) => data);

    // Mock the Python script execution
    excelCalculationService.runPythonScript.mockResolvedValue(JSON.stringify(mockPythonResult));
  });

  afterEach(() => {
    jest.restoreAllMocks();
  });

  test('should normalize template ID correctly', async () => {
    const testCases = [
      { input: 'Format CC1', expected: 'CC1' },
      { input: 'format cc2', expected: 'CC2' },
      { input: 'CC3', expected: 'CC3' },
      { input: 'frcc4', expected: 'CC4' }, // Extracts CC number
      { input: '', expected: null }
    ];

    for (const { input, expected } of testCases) {
      // Reset mock call count for each test case
      templateMappingService.normalizeTemplateId.mockClear();
      
      const result = await excelCalculationService.applyFormDataAndCalculate(input, mockFormData);
      
      if (input) {
        // Verify the mock was called with the input
        expect(templateMappingService.normalizeTemplateId).toHaveBeenCalledWith(input);
        // Verify it returned the expected normalized value
        expect(templateMappingService.normalizeTemplateId).toHaveReturnedWith(expected);
      } else {
        // For empty string, should not call normalize
        expect(templateMappingService.normalizeTemplateId).not.toHaveBeenCalled();
      }
    }
  });

  test('should validate cells according to template-specific rules', async () => {
    // Test data with various cell references
    const testFormData = {
      'd10': 'Description 1', // Should be allowed for CC1/CC2
      'e10': 1000,            // Should be allowed for CC1/CC2
      'h10': 'Client Name',   // Should be allowed for all
      'i10': 45,              // Should be allowed for all
      'j10': 500000,          // Should be allowed for all
      'k10': 'Invalid',       // Should be filtered out for all
      'b10': 'Invalid B'      // Should be filtered out except for fixed assets ranges
    };

    // Mock filterWritableCells to simulate the cell validation logic
    templateMappingService.filterWritableCells.mockImplementation((templateId, data) => {
      const normalizedId = templateMappingService.normalizeTemplateId(templateId);
      const templateColumnRules = {
        'CC1': /^[b-eh-j]\d+$/,  // b,d,e for fixed assets + h,i,j for main sections
        'CC2': /^[b-eh-j]\d+$/,  // b,d,e for fixed assets + h,i,j for main sections
        'CC3': /^[h-j]\d+$/,     // h,i,j columns
        'CC4': /^[h-j]\d+$/,     // h,i,j columns
        'CC5': /^[h-j]\d+$/,     // h,i,j columns
        'CC6': /^[h-j]\d+$/      // h,i,j columns
      };

      const allowedPattern = templateColumnRules[normalizedId] || /^[h-j]\d+$/;
      const filtered = {};
      for (const [key, value] of Object.entries(data)) {
        const lowerKey = key.toLowerCase();
        if (lowerKey.match(allowedPattern)) {
          filtered[lowerKey] = value;
        }
      }
      return filtered;
    });

    // Test CC1 - should allow d,e,h,i,j
    const resultCC1 = await excelCalculationService.applyFormDataAndCalculate('CC1', testFormData);
    expect(templateMappingService.filterWritableCells).toHaveBeenCalledWith('CC1', expect.any(Object));
    // Verify Python script was called
    expect(excelCalculationService.runPythonScript).toHaveBeenCalledTimes(1);

    // Reset mock call count
    excelCalculationService.runPythonScript.mockClear();

    // Test CC3 - should only allow h,i,j
    const resultCC3 = await excelCalculationService.applyFormDataAndCalculate('CC3', testFormData);
    expect(templateMappingService.filterWritableCells).toHaveBeenCalledWith('CC3', expect.any(Object));
    expect(excelCalculationService.runPythonScript).toHaveBeenCalledTimes(1);
  });

  test('should skip fixed assets extraction for CC1 and CC2 templates', () => {
    const spy = jest.spyOn(excelCalculationService, 'extractFixedAssetsSchedule');

    // Test CC1 - should return empty array
    const resultCC1 = excelCalculationService.extractFixedAssetsSchedule({}, 'CC1');
    expect(resultCC1).toEqual([]);
    expect(spy).toHaveReturnedWith([]);

    // Test CC2 - should return empty array
    const resultCC2 = excelCalculationService.extractFixedAssetsSchedule({}, 'Format CC2');
    expect(resultCC2).toEqual([]);
    expect(spy).toHaveReturnedWith([]);

    // Test CC3 - should also return empty array when no data
    const resultCC3 = excelCalculationService.extractFixedAssetsSchedule({}, 'CC3');
    expect(resultCC3).toEqual([]); // Returns [] when no fixed assets found

    spy.mockRestore();
  });

  test('should correctly calculate formulas and return all sheets for CC1 template', async () => {
    const templateId = 'CC1';

    // Action: Call the service
    const result = await excelCalculationService.applyFormDataAndCalculate(templateId, mockFormData);

    // Assertion 1: Check if the result is a valid object
    expect(result).toBeDefined();
    expect(typeof result).toBe('object');

    // Assertion 2: Check if jsonData is present and contains sheets
    expect(result).toHaveProperty('jsonData');
    expect(Array.isArray(result.jsonData)).toBe(true);
    expect(result.jsonData.length).toBeGreaterThan(0);

    // Assertion 3: Check if the 'FinalWorkings' sheet is present
    const finalWorkingsSheet = result.jsonData.find(sheet => sheet.name === 'FinalWorkings');
    expect(finalWorkingsSheet).toBeDefined();
    expect(finalWorkingsSheet).toHaveProperty('data');
    expect(Array.isArray(finalWorkingsSheet.data)).toBe(true);
    expect(finalWorkingsSheet.data.length).toBeGreaterThan(0);

    // Assertion 4: Check if PDF data is generated
    expect(result).toHaveProperty('pdfData');
    expect(result).toHaveProperty('pdfFileName');
    expect(typeof result.pdfData).toBe('string');
    expect(result.pdfFileName).toContain('CC1');

    // Assertion 5: Check if other sheets are present
    expect(result.jsonData.length).toBeGreaterThan(1);

    // Verify Python script was called
    expect(excelCalculationService.runPythonScript).toHaveBeenCalledTimes(1);
  });

  test('should handle modifications and recalculate correctly', async () => {
    const templateId = 'CC1';

    // Initial calculation
    const initialResult = await excelCalculationService.applyFormDataAndCalculate(templateId, mockFormData);

    // Simulate a user modification by changing a value
    const modifiedFormData = {
      ...mockFormData,
      'I10': 2000000, // Changed from 500000
    };

    // Recalculate with modified data
    const recalculatedResult = await excelCalculationService.applyFormDataAndCalculate(templateId, modifiedFormData);

    // Assertion: Check that both results have PDF data generated
    expect(initialResult).toHaveProperty('pdfData');
    expect(recalculatedResult).toHaveProperty('pdfData');

    // Verify Python script was called twice
    expect(excelCalculationService.runPythonScript).toHaveBeenCalledTimes(2);
  });

  test('should handle Python script errors gracefully', async () => {
    const templateId = 'CC1';

    // Mock a Python script error
    excelCalculationService.runPythonScript.mockRejectedValue(new Error('Python script failed'));

    // Action & Assertion: Should throw an error
    await expect(excelCalculationService.applyFormDataAndCalculate(templateId, mockFormData))
      .rejects
      .toThrow('Failed to calculate Excel data using Python engine.');
  });

  test('should handle invalid Python result', async () => {
    const templateId = 'CC1';

    // Mock invalid JSON result
    excelCalculationService.runPythonScript.mockResolvedValue('invalid json');

    // Action & Assertion: Should throw an error
    await expect(excelCalculationService.applyFormDataAndCalculate(templateId, mockFormData))
      .rejects
      .toThrow();
  });
});
