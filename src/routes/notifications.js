const express = require('express');
const router = express.Router();
const Notification = require('../models/Notification');
const { verifyToken } = require('../middleware/auth');

// ============================================================================
// NOTIFICATION ROUTES
// ============================================================================

/**
 * Get user's notifications
 * GET /notifications
 */
router.get('/', verifyToken, async (req, res) => {
  try {
    const { page = 1, limit = 20, unread_only = false } = req.query;
    const skip = (parseInt(page) - 1) * parseInt(limit);

    const filter = { user_id: req.user._id };
    if (unread_only === 'true') {
      filter.read = false;
    }

    const [notifications, total, unreadCount] = await Promise.all([
      Notification.find(filter)
        .sort('-createdAt')
        .skip(skip)
        .limit(parseInt(limit))
        .lean(),
      Notification.countDocuments(filter),
      Notification.countDocuments({ user_id: req.user._id, read: false })
    ]);

    res.json({
      success: true,
      data: {
        notifications,
        unread_count: unreadCount,
        pagination: {
          current_page: parseInt(page),
          total_pages: Math.ceil(total / parseInt(limit)),
          total_count: total
        }
      }
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

/**
 * Get unread notification count
 * GET /notifications/unread-count
 */
router.get('/unread-count', verifyToken, async (req, res) => {
  try {
    const count = await Notification.getUnreadCount(req.user._id);
    res.json({
      success: true,
      data: { count }
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

/**
 * Get recent notifications (for dropdown)
 * GET /notifications/recent
 */
router.get('/recent', verifyToken, async (req, res) => {
  try {
    const notifications = await Notification.find({ user_id: req.user._id })
      .sort('-createdAt')
      .limit(10)
      .lean();

    const unreadCount = await Notification.getUnreadCount(req.user._id);

    res.json({
      success: true,
      data: {
        notifications,
        unread_count: unreadCount
      }
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

/**
 * Mark single notification as read
 * PATCH /notifications/:id/read
 */
router.patch('/:id/read', verifyToken, async (req, res) => {
  try {
    const notification = await Notification.findOneAndUpdate(
      { _id: req.params.id, user_id: req.user._id },
      { $set: { read: true, read_at: new Date() } },
      { new: true }
    );

    if (!notification) {
      return res.status(404).json({ error: 'Notification not found' });
    }

    res.json({
      success: true,
      data: notification
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

/**
 * Mark all notifications as read
 * PATCH /notifications/read-all
 */
router.patch('/read-all', verifyToken, async (req, res) => {
  try {
    const count = await Notification.markAllAsRead(req.user._id);
    
    res.json({
      success: true,
      message: `${count} notifications marked as read`
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

/**
 * Delete a notification
 * DELETE /notifications/:id
 */
router.delete('/:id', verifyToken, async (req, res) => {
  try {
    const notification = await Notification.findOneAndDelete({
      _id: req.params.id,
      user_id: req.user._id
    });

    if (!notification) {
      return res.status(404).json({ error: 'Notification not found' });
    }

    res.json({
      success: true,
      message: 'Notification deleted'
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

/**
 * Clear all notifications
 * DELETE /notifications/clear-all
 */
router.delete('/clear-all', verifyToken, async (req, res) => {
  try {
    const result = await Notification.deleteMany({ user_id: req.user._id });
    
    res.json({
      success: true,
      message: `${result.deletedCount} notifications cleared`
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;
