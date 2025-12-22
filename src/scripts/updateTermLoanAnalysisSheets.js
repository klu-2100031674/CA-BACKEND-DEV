/**
 * Script to update Term Loan templates with Analysis Sheets configuration
 * Usage: node src/scripts/updateTermLoanAnalysisSheets.js
 */

require('dotenv').config({ path: require('path').join(__dirname, '../../.env') });
const mongoose = require('mongoose');
const TemplateConfig = require('../models/TemplateConfig');
const config = require('../config/environment');

const ANALYSIS_SHEETS = [
  { sheet_name: "CFs", display_name: "Cash Flow Statement for the Project", required: true, amount_display: true },
  { sheet_name: "IRR", display_name: "IRR for the Project", required: true, amount_display: true },
  { sheet_name: "MIRR", display_name: "Modified IRR for the Project", required: true, amount_display: true },
  { sheet_name: "Funds Flow statement", display_name: "Funds Flow Statement for the Project", required: true, amount_display: true },
  { sheet_name: "index", display_name: "Profitability Index for the Project", required: true, amount_display: true },
  { sheet_name: "NVP", display_name: "Net Present Value (NPV) for the Project", required: true, amount_display: true },
  { sheet_name: "Payback period 1", display_name: "Investment Payback Period for the Project (1)", required: true, amount_display: true },
  { sheet_name: "Payback period 2", display_name: "Investment Payback Period for the Project (2)", required: true, amount_display: true },
  { sheet_name: "Sensitivity Analysis", display_name: "Sensitivity Analysis for the Project", required: true, amount_display: true },
  { sheet_name: "RATIO", display_name: "Leverage Ratio Analysis for the Project", required: true, amount_display: true },
  { sheet_name: "Altman Z", display_name: "Altman Z Score of the Project", required: true, amount_display: true },
  { sheet_name: "BEP", display_name: "Breakeven Point (BEP) for the Project", required: true, amount_display: true }
];

async function updateTemplates() {
  try {
    console.log('🚀 Connecting to MongoDB...');
    await mongoose.connect(config.MONGODB_URI);
    console.log('✅ Connected to MongoDB');

    const termLoanTemplates = await TemplateConfig.find({
      $or: [
        { template_id: /TERM_LOAN/i },
        { report_type: 'TERM_LOAN' }
      ]
    });

    console.log(`📄 Found ${termLoanTemplates.length} Term Loan templates to update`);

    for (const template of termLoanTemplates) {
      template.analysis_sheets = ANALYSIS_SHEETS;
      await template.save();
      console.log(`✅ Updated: ${template.name} (${template.template_id})`);
    }

    console.log('\n✨ All Term Loan templates updated successfully!');
    process.exit(0);
  } catch (error) {
    console.error('❌ Error:', error.message);
    process.exit(1);
  }
}

updateTemplates();
