/**
 * Report Queue Service
 * Handles serialization of Excel COM operations to prevent concurrent access issues.
 * Excel COM automation is NOT thread-safe - only one Excel operation should run at a time.
 */

const logger = require('../utils/logger');

class ReportQueueService {
  constructor() {
    this.queue = [];
    this.activeWorkers = 0;
    this.maxConcurrency = 4; // Allow up to 4 concurrent Excel operations
    this.maxQueueSize = 100; // Increased pending reports limit
    this.taskTimeout = 10 * 60 * 1000; // 10 minutes timeout per task (Excel can be slow)
  }

  /**
   * Add a report generation task to the queue
   * @param {Function} task - Async function to execute
   * @param {Object} metadata - Task metadata for logging
   * @returns {Promise} - Resolves with task result
   */
  enqueue(task, metadata = {}) {
    return new Promise((resolve, reject) => {
      // Check queue size limit
      if (this.queue.length >= this.maxQueueSize) {
        const error = new Error('Report generation queue is full. Please try again later.');
        error.statusCode = 503;
        logger.warn('Report queue full', {
          operation: 'reportQueue.enqueue',
          queueSize: this.queue.length,
          maxSize: this.maxQueueSize,
          ...metadata
        });
        return reject(error);
      }

      const queueItem = {
        id: `task_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
        task,
        metadata,
        resolve,
        reject,
        enqueuedAt: Date.now()
      };

      this.queue.push(queueItem);
      
      logger.info('Report task enqueued', {
        operation: 'reportQueue.enqueue',
        taskId: queueItem.id,
        queuePosition: this.queue.length,
        ...metadata
      });

      // Start processing if not already running
      this.processQueue();
    });
  }

  /**
   * Process the queue - allows multiple concurrent tasks up to maxConcurrency
   */
  async processQueue() {
    // If we've reached max concurrency or queue is empty, nothing to do
    if (this.activeWorkers >= this.maxConcurrency || this.queue.length === 0) {
      return;
    }

    // Launch workers up to maxConcurrency
    while (this.activeWorkers < this.maxConcurrency && this.queue.length > 0) {
      const item = this.queue.shift();
      this.activeWorkers++;

      logger.info('Processing report task', {
        operation: 'reportQueue.process',
        taskId: item.id,
        activeWorkers: this.activeWorkers,
        waitTime: Date.now() - item.enqueuedAt,
        remainingInQueue: this.queue.length,
        ...item.metadata
      });

      // Run task in background (not awaited here so we can start more tasks)
      this.runTask(item);
    }
  }

  /**
   * Internal helper to run a specific task and manage worker count
   */
  async runTask(item) {
    try {
      // Execute the task with timeout
      const result = await this.executeWithTimeout(item.task, item.id);
      
      logger.info('Report task completed', {
        operation: 'reportQueue.complete',
        taskId: item.id,
        duration: Date.now() - item.enqueuedAt,
        ...item.metadata
      });

      item.resolve(result);
    } catch (error) {
      logger.error('Report task failed', {
        operation: 'reportQueue.error',
        taskId: item.id,
        error: error.message,
        stack: error.stack,
        ...item.metadata
      });

      item.reject(error);
    } finally {
      this.activeWorkers--;
      // Try to process next items in queue
      this.processQueue();
    }
  }

  /**
   * Execute a task with timeout
   */
  async executeWithTimeout(task, taskId) {
    return new Promise((resolve, reject) => {
      const timeoutId = setTimeout(() => {
        reject(new Error(`Report generation timeout after ${this.taskTimeout / 1000} seconds`));
      }, this.taskTimeout);

      task()
        .then((result) => {
          clearTimeout(timeoutId);
          resolve(result);
        })
        .catch((error) => {
          clearTimeout(timeoutId);
          reject(error);
        });
    });
  }

  /**
   * Get current queue status
   */
  getStatus() {
    return {
      queueLength: this.queue.length,
      activeWorkers: this.activeWorkers,
      maxConcurrency: this.maxConcurrency,
      pendingTasks: this.queue.map((item, index) => ({
        position: index + 1,
        id: item.id,
        enqueuedAt: item.enqueuedAt,
        metadata: item.metadata
      }))
    };
  }

  /**
   * Get estimated wait time for a new task
   */
  getEstimatedWaitTime() {
    // Assume average task takes 45 seconds (increased for concurrent load)
    const avgTaskTime = 45 * 1000;
    const totalPending = this.queue.length;
    
    // Estimated time is (pending tasks / concurrency) * avg time
    const rounds = Math.ceil((totalPending + this.activeWorkers) / this.maxConcurrency);
    return rounds * avgTaskTime;
  }
}

// Singleton instance
const reportQueueService = new ReportQueueService();

module.exports = reportQueueService;
