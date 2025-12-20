const mongoose = require('mongoose');
const { Schema } = mongoose;

/**
 * Notification Schema
 * Stores notifications for users (approval, rejection, system messages, etc.)
 */
const notificationSchema = new Schema({
  user_id: { 
    type: Schema.Types.ObjectId, 
    ref: 'User', 
    required: true,
    index: true
  },
  
  type: { 
    type: String, 
    enum: [
      'report_approved',
      'report_rejected', 
      'report_under_review',
      'report_revised',
      'payment_received',
      'payment_failed',
      'withdrawal_approved',
      'withdrawal_rejected',
      'commission_earned',
      'system_message',
      'welcome'
    ],
    required: true
  },
  
  title: { type: String, required: true },
  message: { type: String, required: true },
  
  // Reference data (optional - for linking to specific resources)
  data: {
    report_id: { type: Schema.Types.ObjectId, ref: 'Report' },
    withdrawal_id: { type: Schema.Types.ObjectId, ref: 'Withdrawal' },
    amount: { type: Number },
    extra: { type: Schema.Types.Mixed }
  },
  
  read: { type: Boolean, default: false },
  read_at: { type: Date },
  
  // For email tracking
  email_sent: { type: Boolean, default: false },
  email_sent_at: { type: Date }
}, { 
  timestamps: true 
});

// Indexes for efficient queries
notificationSchema.index({ user_id: 1, read: 1, createdAt: -1 });
notificationSchema.index({ createdAt: -1 });

// Static method to create and optionally send notification
notificationSchema.statics.createNotification = async function(data) {
  const notification = new this(data);
  await notification.save();
  return notification;
};

// Static method to get unread count for a user
notificationSchema.statics.getUnreadCount = async function(userId) {
  return this.countDocuments({ user_id: userId, read: false });
};

// Static method to mark all as read for a user
notificationSchema.statics.markAllAsRead = async function(userId) {
  const result = await this.updateMany(
    { user_id: userId, read: false },
    { $set: { read: true, read_at: new Date() } }
  );
  return result.modifiedCount;
};

module.exports = mongoose.model('Notification', notificationSchema);
