const mongoose = require('mongoose');
const Report = require('./src/models/Report');
require('dotenv').config();

async function checkReports() {
  try {
    await mongoose.connect(process.env.MONGODB_URI || 'mongodb://127.0.0.1:27017/ca_report_generation');
    
    const reports = await Report.find({})
      .select('title excel_file_url pdf_file_url validation_status createdAt')
      .sort({ createdAt: -1 })
      .limit(3);
    
    console.log(`\n📊 Last 3 Reports:\n`);
    
    reports.forEach((report, index) => {
      console.log(`${index + 1}. ${report.title || 'Untitled'}`);
      console.log(`   ID: ${report._id}`);
      console.log(`   Created: ${report.createdAt}`);
      console.log(`   Status: ${report.validation_status}`);
      console.log(`   Excel URL: ${report.excel_file_url || '❌ NOT SET'}`);
      console.log(`   PDF URL: ${report.pdf_file_url || '❌ NOT SET'}`);
      
      if (report.excel_file_url && report.excel_file_url.includes('r2.cloudflarestorage.com')) {
        console.log('   ✅ Using R2 Cloud Storage');
      } else {
        console.log('   ⚠️  NOT using R2 - Local path or missing');
      }
      console.log('');
    });
    
  } catch (error) {
    console.error('Error:', error.message);
  } finally {
    await mongoose.connection.close();
    process.exit(0);
  }
}

checkReports();
