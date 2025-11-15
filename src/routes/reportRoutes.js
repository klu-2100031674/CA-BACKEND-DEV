const express = require('express');
const router = express.Router();
const reportController = require('../controllers/reportController');
const { verifyToken } = require('../middleware/auth');

/**
 * Report Routes
 */

// Apply form data to template and get calculated Excel
router.post(
  '/templates/:templateId/apply-form',
  verifyToken,
  reportController.applyFormData
);

// Apply final edits (sheet updates) and get recalculated sheets
router.post(
  '/templates/:templateId/apply-final',
  verifyToken,
  reportController.applyFinalEdits
);

// Download calculated Excel file
router.get(
  '/templates/:templateId/download/:fileName',
  verifyToken,
  reportController.downloadExcelFile
);

// Export to PDF
router.post(
  '/export/pdf',
  verifyToken,
  reportController.exportToPdf
);

// Export to JSON
router.post(
  '/export/json',
  verifyToken,
  reportController.exportToJson
);

// Download full AI-enhanced report (authentication optional for testing)
router.post(
  '/templates/:templateId/download-full-report',
  reportController.downloadFullReport
);

// Download exported file
router.get(
  '/download/:fileName',
  verifyToken,
  reportController.downloadExcelFile
);

// CRUD operations
router.post('/', verifyToken, reportController.createReport);
router.get('/', verifyToken, reportController.getReports);
router.get('/:reportId', verifyToken, reportController.getReportById);
router.delete('/:reportId', verifyToken, reportController.deleteReport);

// Upload JSON data (for backwards compatibility)
router.post('/:reportId/upload-json', reportController.uploadReportJson);

module.exports = router;
