const express = require('express');
const router = express.Router();
const excelFileService = require('../services/excelFileService');
const { verifyToken } = require('../middleware/auth');
const logger = require('../utils/logger');

/**
 * Excel Files Routes
 * Handles Excel file storage and retrieval from MongoDB
 */

/**
 * @route   GET /api/excel-files
 * @desc    Get user's Excel files
 * @access  Private
 */
router.get('/', verifyToken, async (req, res, next) => {
  try {
    const { page, limit, stage, templateId } = req.query;

    const result = await excelFileService.getExcelFilesByUser(req.user._id, {
      page: parseInt(page) || 1,
      limit: parseInt(limit) || 10,
      stage,
      templateId
    });

    res.json({
      success: true,
      data: result
    });

  } catch (error) {
    logger.error('Error in getExcelFiles', {
      userId: req.user._id,
      error: error.message,
      stack: error.stack,
      operation: 'getExcelFiles'
    });
    next(error);
  }
});

/**
 * @route   GET /api/excel-files/:fileId
 * @desc    Get Excel file details
 * @access  Private
 */
router.get('/:fileId', verifyToken, async (req, res, next) => {
  try {
    const file = await excelFileService.getExcelFile(req.params.fileId);

    // Check ownership
    if (file.user_id._id.toString() !== req.user._id.toString() && req.user.role !== 'admin' && req.user.role !== 'super_admin') {
      return res.status(403).json({
        success: false,
        error: 'Access denied'
      });
    }

    res.json({
      success: true,
      data: file
    });

  } catch (error) {
    logger.error('Error in getExcelFile', {
      userId: req.user._id,
      fileId: req.params.fileId,
      error: error.message,
      stack: error.stack,
      operation: 'getExcelFile'
    });
    if (error.message.includes('not found')) {
      return res.status(404).json({
        success: false,
        error: 'File not found'
      });
    }
    next(error);
  }
});

/**
 * @route   DELETE /api/excel-files/:fileId
 * @desc    Delete Excel file
 * @access  Private
 */
router.delete('/:fileId', verifyToken, async (req, res, next) => {
  try {
    const success = await excelFileService.deleteExcelFile(req.params.fileId, req.user._id);

    res.json({
      success: true,
      message: 'File deleted successfully'
    });

  } catch (error) {
    logger.error('Error in deleteExcelFile', {
      userId: req.user._id,
      fileId: req.params.fileId,
      error: error.message,
      stack: error.stack,
      operation: 'deleteExcelFile'
    });
    if (error.message.includes('not found')) {
      return res.status(404).json({
        success: false,
        error: 'File not found'
      });
    }
    if (error.message.includes('access denied')) {
      return res.status(403).json({
        success: false,
        error: 'Access denied'
      });
    }
    next(error);
  }
});

module.exports = router;