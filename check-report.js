const mongoose = require('mongoose');
const Report = require('./src/models/Report');
require('dotenv').config();

async function checkReport() {
  try {
    await mongoose.connect(process.env.MONGODB_URI || 'mongodb://127.0.0.1:27017/ca_report_generation');
    
    const reportId = '693eaf29a21c42385e7fe2b4';
    const report = await Report.findById(reportId).select('title excel_file_url pdf_file_url validation_status createdAt user_id');
    
    if (!report) {
      console.log('❌ Report not found!');
      process.exit(1);
    }
    
    console.log('\n📄 Report Details:\n');
    console.log('Title:', report.title);
    console.log('Status:', report.validation_status);
    console.log('Created:', report.createdAt);
    console.log('\n📁 File URLs:\n');
    console.log('Excel URL:', report.excel_file_url || 'NOT SET');
    console.log('PDF URL:', report.pdf_file_url || 'NOT SET');
    
    console.log('\n🔍 Analysis:\n');
    
    if (report.excel_file_url) {
      if (report.excel_file_url.includes('r2.cloudflarestorage.com')) {
        console.log('✅ Excel: R2 Cloud Storage URL');
      } else {
        console.log('⚠️  Excel: Local file path (legacy)');
      }
    } else {
      console.log('❌ Excel: NO URL');
    }
    
    if (report.pdf_file_url) {
      if (report.pdf_file_url.includes('r2.cloudflarestorage.com')) {
        console.log('✅ PDF: R2 Cloud Storage URL');
      } else {
        console.log('⚠️  PDF: Local file path (legacy)');
      }
    } else {
      console.log('❌ PDF: NO URL');
    }
    
    console.log('\n');
    
  } catch (error) {
    console.error('Error:', error.message);
  } finally {
    await mongoose.connection.close();
    process.exit(0);
  }
}

checkReport();
