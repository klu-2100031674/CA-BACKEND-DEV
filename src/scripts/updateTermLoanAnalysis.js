/**
 * Script to update Term Loan templates with analysis sheet configurations
 * Usage: node src/scripts/updateTermLoanAnalysis.js
 */

require('dotenv').config();
const mongoose = require('mongoose');
const TemplateConfig = require('../models/TemplateConfig');
const config = require('../config/environment');

const ANALYSIS_SHEETS = [
  { sheet_name: 'CFs', display_name: 'Cash Flow Statement', required: true, price: 500, amount_display: '500' },
  { sheet_name: 'IRR', display_name: 'IRR Analysis', required: true, price: 500, amount_display: '500' },
  { sheet_name: 'BEP analysis', display_name: 'Break Even Point (BEP)', required: true, price: 500, amount_display: '₹500' },
  { sheet_name: 'RATIO', display_name: 'Ratio Analysis', required: true, price: 0, amount_display: 'Included' },
  { sheet_name: 'Sensitivity Analysis', display_name: 'Sensitivity Analysis', required: true, price: 500, amount_display: '₹500' },
  { sheet_name: 'DSCR', display_name: 'DSCR Analysis', required: false, price: 0, amount_display: 'Included' },
  { sheet_name: 'NPV', display_name: 'NPV Analysis', required: false, price: 0, amount_display: 'Included' },
  { sheet_name: 'WACC', display_name: 'WACC Analysis', required: false, price: 0, amount_display: 'Included' }
];

async function updateTemplates() {
  try {
    console.log('🚀 Connecting to MongoDB...');
    await mongoose.connect(config.MONGODB_URI);
    console.log('✅ Connected');

    const termLoanTemplates = [
      'TERM_LOAN_SERVICE_WITHOUT_STOCK',
      'TERM_LOAN_MANUFACTURING_SERVICE_WITH_STOCK',
      'TERM_LOAN_CC'
    ];

    for (const tid of termLoanTemplates) {
      const template = await TemplateConfig.findOne({ template_id: tid });
      if (template) {
        template.analysis_sheets = ANALYSIS_SHEETS;
        await template.save();
        console.log(`✅ Updated analysis sheets for: ${tid}`);
      } else {
        console.log(`⚠️ Template not found: ${tid}`);
      }
    }

    console.log('✨ All Term Loan templates updated successfully');
    process.exit(0);
  } catch (error) {
    console.error('❌ Error:', error);
    process.exit(1);
  }
}

updateTemplates();
