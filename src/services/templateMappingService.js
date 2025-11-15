/**
 * Template Mapping Service
 * 
 * Provides dynamic cell mapping based on Excel template structure.
 * Each CC format has different row/column layouts - this service ensures
 * we only write to INPUT cells and never overwrite FORMULA cells.
 */

const fs = require('fs');
const path = require('path');

class TemplateMappingService {
  constructor() {
    this.mappingsPath = path.join(__dirname, '../../templates/excel/template-mappings.json');
    this.mappings = null;
    this.loadMappings();
  }

  /**
   * Load template mappings from JSON file
   */
  loadMappings() {
    try {
      const data = fs.readFileSync(this.mappingsPath, 'utf8');
      this.mappings = JSON.parse(data);
      console.log('[TemplateMappingService] Loaded mappings for:', Object.keys(this.mappings).join(', '));
    } catch (error) {
      console.error('[TemplateMappingService] Failed to load mappings:', error.message);
      this.mappings = {};
    }
  }

  /**
   * Get mapping for a specific template
   * @param {string} templateId - Template ID (e.g., 'frcc2', 'Format CC2', 'CC2')
   * @returns {object} Template mapping configuration
   */
  getMapping(templateId) {
    // Normalize template ID to CC format
    const normalizedId = this.normalizeTemplateId(templateId);
    
    if (!this.mappings[normalizedId]) {
      console.warn(`[TemplateMappingService] No mapping found for ${templateId} (normalized: ${normalizedId})`);
      return null;
    }

    return this.mappings[normalizedId];
  }

  /**
   * Normalize template ID to standard CC format
   * @param {string} templateId - Template ID in any format
   * @returns {string} Normalized ID (e.g., 'CC2')
   */
  normalizeTemplateId(templateId) {
    if (!templateId) return null;
    
    const idUpper = templateId.toUpperCase();
    
    // Extract CC number
    const match = idUpper.match(/CC(\d+)/);
    if (match) {
      return `CC${match[1]}`;
    }
    
    return templateId;
  }

  /**
   * Get input rows for a specific section
   * @param {string} templateId - Template ID
   * @param {string} section - Section name ('audited', 'provisional', 'assumptions')
   * @returns {array} Array of row numbers that are safe to write
   */
  getInputRows(templateId, section) {
    const mapping = this.getMapping(templateId);
    if (!mapping) return [];

    const sectionKey = `${section}_section`;
    if (!mapping[sectionKey] || !mapping[sectionKey].input_rows) {
      console.warn(`[TemplateMappingService] No input rows found for ${templateId}.${section}`);
      return [];
    }

    return mapping[sectionKey].input_rows;
  }

  /**
   * Get formula rows for a specific section (rows to SKIP)
   * @param {string} templateId - Template ID
   * @param {string} section - Section name ('audited', 'provisional', 'assumptions')
   * @returns {array} Array of row numbers that contain formulas (DO NOT WRITE)
   */
  getFormulaRows(templateId, section) {
    const mapping = this.getMapping(templateId);
    if (!mapping) return [];

    const sectionKey = `${section}_section`;
    if (!mapping[sectionKey] || !mapping[sectionKey].formula_rows) {
      return [];
    }

    return mapping[sectionKey].formula_rows;
  }

  /**
   * Get finance section cell types (H vs I columns)
   * @param {string} templateId - Template ID
   * @returns {object} Object mapping cell references to types ('input' or 'formula')
   */
  getFinanceCells(templateId) {
    const mapping = this.getMapping(templateId);
    if (!mapping || !mapping.finance_section || !mapping.finance_section.rows) {
      return {};
    }

    return mapping.finance_section.rows;
  }

  /**
   * Get fixed assets row mappings
   * @param {string} templateId - Template ID
   * @returns {object} Object mapping category names to row numbers
   */
  getFixedAssetsMapping(templateId) {
    const mapping = this.getMapping(templateId);
    if (!mapping || !mapping.fixed_assets) {
      console.warn(`[TemplateMappingService] No fixed assets mapping for ${templateId}`);
      return {};
    }

    return mapping.fixed_assets;
  }

  /**
   * Check if a cell should be written (not a formula cell)
   * @param {string} templateId - Template ID
   * @param {string} section - Section name
   * @param {number} rowNumber - Row number
   * @returns {boolean} True if safe to write, false if formula cell
   */
  canWriteToRow(templateId, section, rowNumber) {
    const inputRows = this.getInputRows(templateId, section);
    const formulaRows = this.getFormulaRows(templateId, section);

    // If it's a formula row, definitely can't write
    if (formulaRows.includes(rowNumber)) {
      return false;
    }

    // If it's in the input rows list, can write
    if (inputRows.includes(rowNumber)) {
      return true;
    }

    // If not in either list, be conservative and don't write
    console.warn(`[TemplateMappingService] Row ${rowNumber} not in mapping for ${templateId}.${section} - skipping`);
    return false;
  }

  /**
   * Filter cell data to only include writable cells
   * @param {string} templateId - Template ID
   * @param {object} cellData - Object with cell references as keys
   * @returns {object} Filtered cell data with only writable cells
   */
  filterWritableCells(templateId, cellData) {
    const mapping = this.getMapping(templateId);
    if (!mapping) {
      console.warn(`[TemplateMappingService] No mapping for ${templateId}, allowing all cells`);
      return cellData;
    }

    const filtered = {};
    const financeCells = this.getFinanceCells(templateId);

    for (const [cellRef, value] of Object.entries(cellData)) {
      const lowerRef = cellRef.toLowerCase();

      // Check finance cells (h12, h13, etc.)
      if (financeCells[lowerRef]) {
        if (financeCells[lowerRef] === 'input') {
          filtered[cellRef] = value;
        } else {
          console.log(`[TemplateMappingService] Skipping formula cell: ${cellRef}`);
        }
        continue;
      }

      // Extract row number from cell reference (e.g., 'i23' -> 23)
      const rowMatch = lowerRef.match(/[a-z](\d+)/);
      if (!rowMatch) continue;

      const rowNum = parseInt(rowMatch[1]);

      // Determine section based on row number (rough heuristic)
      let section = null;
      if (rowNum >= 22 && rowNum <= 43) section = 'audited';
      else if (rowNum >= 44 && rowNum <= 66) section = 'provisional';
      else if (rowNum >= 67 && rowNum <= 99) section = 'assumptions';
      else if ((templateId === 'CC1' || templateId === 'Format CC1') && rowNum >= 100 && rowNum <= 200) {
        // Allow fixed assets cells for CC1 (b100-e200 range)
        filtered[cellRef] = value;
        continue;
      }
      else if ((templateId === 'CC2' || templateId === 'Format CC2') && rowNum >= 121 && rowNum <= 221) {
        // Allow fixed assets cells for CC2 (b121-e221 range)
        filtered[cellRef] = value;
        continue;
      }

      if (section && this.canWriteToRow(templateId, section, rowNum)) {
        filtered[cellRef] = value;
      } else if (!section) {
        // For cells outside known sections, allow them (general info, etc.)
        filtered[cellRef] = value;
      }
    }

    console.log(`[TemplateMappingService] Filtered ${Object.keys(cellData).length} cells -> ${Object.keys(filtered).length} writable cells`);
    return filtered;
  }
}

// Singleton instance
const templateMappingService = new TemplateMappingService();

module.exports = templateMappingService;
