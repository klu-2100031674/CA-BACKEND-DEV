/**
 * Script to update sheet names and visibility for Service Sector Without Stock template
 * 
 * Usage: node src/scripts/updateServiceSectorSheets.js
 */

require('dotenv').config();
const mongoose = require('mongoose');
const TemplateConfig = require('../models/TemplateConfig');
const config = require('../config/environment');

async function updateSheets() {
  try {
    console.log('🚀 Connecting to MongoDB...');
    await mongoose.connect(config.MONGODB_URI);
    console.log('✅ Connected to MongoDB');

    const serviceWithoutStockSheets = [
      { sheet_name: 'Cover page', display_name: 'Cover page', is_included: true, is_optional: false, is_visible: true },
      { sheet_name: 'PL BS', display_name: 'PL BS', is_included: true, is_optional: false, is_visible: true },
      { sheet_name: 'Assumptions', display_name: 'Assumptions', is_included: true, is_optional: false, is_visible: true },
      { sheet_name: 'Loan Schd', display_name: 'Term Loan', is_included: true, is_optional: false, is_visible: true },
      { sheet_name: 'FA Sch', display_name: 'Depreciation', is_included: true, is_optional: false, is_visible: true },
      { sheet_name: 'Repayment(Principal)', display_name: 'Repayment', is_included: true, is_optional: false, is_visible: true },
      { sheet_name: 'RATIO', display_name: 'Ratio', is_included: true, is_optional: false, is_visible: true },
      { sheet_name: 'CFs', display_name: 'Cash Flow', is_included: true, is_optional: false, is_visible: true },
      { sheet_name: 'Sensitivity Analysis', display_name: 'Sensitivity', is_included: false, is_optional: true, is_visible: true },
      { sheet_name: 'BEP analysis', display_name: 'BEP', is_included: false, is_optional: false, is_visible: false },
      { sheet_name: 'Gaurantors', display_name: 'Gaurantors', is_included: true, is_optional: false, is_visible: true }
    ];

    const manufacturingWithStockSheets = [
      { sheet_name: 'Cover page', display_name: 'Cover page', is_included: true, is_optional: false, is_visible: true },
      { sheet_name: 'PL BS', display_name: 'PL BS', is_included: true, is_optional: false, is_visible: true },
      { sheet_name: 'Assumptions', display_name: 'Assumptions', is_included: true, is_optional: false, is_visible: true },
      { sheet_name: 'Loan sch', display_name: 'Term Loan', is_included: true, is_optional: false, is_visible: true },
      { sheet_name: 'FA Sch', display_name: 'Depreciation', is_included: true, is_optional: false, is_visible: true },
      { sheet_name: 'Repayment', display_name: 'Repayment', is_included: true, is_optional: false, is_visible: true },
      { sheet_name: 'RATIO', display_name: 'Ratio', is_included: true, is_optional: false, is_visible: true },
      { sheet_name: 'CFs', display_name: 'Cash Flow', is_included: true, is_optional: false, is_visible: true },
      { sheet_name: 'Sensitivity Analysis', display_name: 'Sensitivity', is_included: false, is_optional: true, is_visible: true },
      { sheet_name: 'BEP analysis', display_name: 'BEP', is_included: false, is_optional: true, is_visible: true },
      { sheet_name: 'Gaurantors', display_name: 'Gaurantors', is_included: true, is_optional: false, is_visible: true },
      { sheet_name: 'Sheet1', display_name: 'Sheet1', is_included: true, is_optional: false, is_visible: true }
    ];

    const templates = [
      { id: 'TERM_LOAN_SERVICE_WITHOUT_STOCK', sheets: serviceWithoutStockSheets },
      { id: 'TERM_LOAN_MANUFACTURING_SERVICE_WITH_STOCK', sheets: manufacturingWithStockSheets }
    ];

    for (const template of templates) {
      const result = await TemplateConfig.findOneAndUpdate(
        { template_id: template.id },
        { 
          $set: { 
            'pricing.sheet_pricing': template.sheets 
          } 
        },
        { new: true }
      );

      if (result) {
        console.log(`✅ Successfully updated sheets for ${template.id}`);
      } else {
        console.error(`❌ Template ${template.id} not found in database.`);
      }
    }

    await mongoose.disconnect();
    console.log('👋 Disconnected from MongoDB');
  } catch (error) {
    console.error('❌ Error:', error.message);
    process.exit(1);
  }
}

updateSheets();
