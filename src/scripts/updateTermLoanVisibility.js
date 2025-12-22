/**
 * Script to update visibility for Term Loan templates
 * Usage: node src/scripts/updateTermLoanVisibility.js
 */

require('dotenv').config();
const mongoose = require('mongoose');
const TemplateConfig = require('../models/TemplateConfig');
const config = require('../config/environment');

const ALL_TERM_LOAN_SHEETS = [
  "Cover page",
  "cover page",
  "Index",
  "profile",
  "Descriptive",
  "project cost",
  "PL BS",
  "PLBS",
  "Working for PL & BS",
  "Workings for PLBS",
  "Graph",
  "RATIO",
  "FA Sch",
  "Dep IT act",
  "Dep sch",
  "Loan sch",
  "Loan Sch",
  "Repayment",
  "Repayment sch",
  "CFs",
  "CF's",
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
  "Final workings",
  "MPBF ",
  "workings for sensitivity1",
  "Gaurantors",
  "BEP analysis",
  "Sales",
  "Assumptions.1",
  "Assumptions 1",
  "Nayak Committee Recommandation",
  "nayak",
  "Sheet3",
  "Sheet2"
];

const VISIBLE_SHEETS = [
  {
    sheet_name: "Assumptions.1",
    display_name: "Assumptions 1",
    required: true,
    is_visible: true,
    price: 0,
    amount_display: "Included"
  },
  {
    sheet_name: "PL BS",
    display_name: "Working for PL & BS",
    required: true,
    is_visible: true,
    price: 0,
    amount_display: "Included"
  },
  {
    sheet_name: "Final workings",
    display_name: "Workings for PLBS",
    required: true,
    is_visible: true,
    price: 0,
    amount_display: "Included"
  },
  {
    sheet_name: "Sensitivity Analysis",
    display_name: "Sensitivity Analysis",
    required: false,
    is_visible: true,
    price: 0,
    amount_display: "Included"
  },
  {
    sheet_name: "BEP analysis",
    display_name: "BEP Analysis",
    required: false,
    is_visible: true,
    price: 0,
    amount_display: "Included"
  },
  {
    sheet_name: "BEP",
    display_name: "BEP",
    required: false,
    is_visible: true,
    price: 0,
    amount_display: "Included"
  }
];

async function updateTermLoanVisibility() {
  try {
    console.log('🚀 Connecting to MongoDB...');
    await mongoose.connect(process.env.MONGODB_URI || config.MONGODB_URI);
    console.log('✅ Connected to MongoDB');

    const termLoanTemplates = await TemplateConfig.find({
      $or: [
        { report_type: 'TERM_LOAN' },
        { template_id: { $regex: /TERM_LOAN/i } }
      ]
    });

    console.log(`📄 Found ${termLoanTemplates.length} Term Loan templates`);

    for (const template of termLoanTemplates) {
      console.log(`Updating template: ${template.template_id}`);
      
      // Clear existing analysis_sheets to ensure fresh start
      template.analysis_sheets = [];
      
      const newAnalysisSheets = [];

      // Process all sheets to ensure they are in analysis_sheets
      for (const sheetName of ALL_TERM_LOAN_SHEETS) {
        const visibleConfig = VISIBLE_SHEETS.find(v => v.sheet_name === sheetName);

        if (visibleConfig) {
          // This is one of the 3 visible sheets
          newAnalysisSheets.push(visibleConfig);
        } else {
          // This is a hidden required sheet
          newAnalysisSheets.push({
            sheet_name: sheetName,
            display_name: sheetName,
            required: true,
            is_visible: false,
            price: 0,
            amount_display: "Included"
          });
        }
      }

      template.analysis_sheets = newAnalysisSheets;

      // 3. Remove visible sheets from after_generate_hide so they appear in PDF
      if (template.after_generate_hide) {
        const visibleSheetNames = VISIBLE_SHEETS.map(v => v.sheet_name);
        const variations = ["Assumptions 1", "Assumptions.1", "Assumptions", "assumptions"];
        
        template.after_generate_hide = template.after_generate_hide.filter(
          sheet => !visibleSheetNames.includes(sheet) && !variations.includes(sheet)
        );
      }

      // 4. Ensure they are in full_report_sheets
      template.full_report_sheets = ALL_TERM_LOAN_SHEETS;

      template.markModified('analysis_sheets');
      template.markModified('after_generate_hide');
      template.markModified('full_report_sheets');
      
      await template.save();
      console.log(`✅ Updated: ${template.template_id}`);
    }

    console.log('✨ All Term Loan templates updated successfully');
    process.exit(0);
  } catch (error) {
    console.error('❌ Error updating templates:', error);
    process.exit(1);
  }
}

updateTermLoanVisibility();
