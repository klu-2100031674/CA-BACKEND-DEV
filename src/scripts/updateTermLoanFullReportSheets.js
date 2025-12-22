/**
 * Script to update full_report_sheets for Term Loan templates
 * Usage: node src/scripts/updateTermLoanFullReportSheets.js
 */

require('dotenv').config();
const mongoose = require('mongoose');
const TemplateConfig = require('../models/TemplateConfig');
const config = require('../config/environment');

const TERM_LOAN_FULL_REPORT_SHEETS = [
  "Cover page",
  "Index",
  "profile",
  "Descriptive",
  "project cost",
  "PL BS",
  "Graph",
  "RATIO",
  "FA Sch",
  "Dep IT act",
  "Loan sch",
  "Repayment",
  "CFs",
  "IRR",
  "MIRR",
  "NPV",
  "PI Index",
  "WACC",
  "Payback period I",
  "Payback period II",
  "Altman Z",
  "Sensitivity Analysis",
  "workings for sensittivity1",
  "Workings for Sensitivity2",
  "CF workings",
  "Final workings",
  "MPBF ",
  "workings for sensitivity1",
  "Gaurantors",
  "BEP analysis",
  "Sales",
];

async function updateTermLoanSheets() {
  try {
    console.log('🚀 Connecting to MongoDB...');
    await mongoose.connect(config.MONGODB_URI);
    console.log('✅ Connected to MongoDB');

    const termLoanTemplates = await TemplateConfig.find({
      template_id: { $regex: /TERM_LOAN/i }
    });

    console.log(`📄 Found ${termLoanTemplates.length} Term Loan templates`);

    for (const template of termLoanTemplates) {
      template.full_report_sheets = TERM_LOAN_FULL_REPORT_SHEETS;
      await template.save();
      console.log(`✅ Updated sheets for: ${template.template_id}`);
    }

    console.log('✨ All Term Loan templates updated successfully');
    process.exit(0);
  } catch (error) {
    console.error('❌ Error updating templates:', error);
    process.exit(1);
  }
}

updateTermLoanSheets();
