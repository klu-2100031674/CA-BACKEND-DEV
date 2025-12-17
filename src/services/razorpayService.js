const Razorpay = require('razorpay');
const crypto = require('crypto');
const https = require('https');
const logger = require('../utils/logger');
const config = require('../config/environment');

class RazorpayService {
  constructor() {
    this.razorpay = new Razorpay({
      key_id: config.RAZORPAY_KEY_ID,
      key_secret: config.RAZORPAY_KEY_SECRET,
    });
    this.isProduction = config.NODE_ENV === 'production';
  }

  /**
   * Make HTTP request to Razorpay API
   * @param {string} endpoint - API endpoint
   * @param {string} method - HTTP method
   * @param {Object} data - Request data
   * @returns {Promise<Object>} API response
   */
  async makeRazorpayRequest(endpoint, method = 'POST', data = null) {
    return new Promise((resolve, reject) => {
      const auth = Buffer.from(`${config.RAZORPAY_KEY_ID}:${config.RAZORPAY_KEY_SECRET}`).toString('base64');

      const options = {
        hostname: 'api.razorpay.com',
        port: 443,
        path: `/v1${endpoint}`,
        method: method,
        headers: {
          'Authorization': `Basic ${auth}`,
          'Content-Type': 'application/json',
        }
      };

      const req = https.request(options, (res) => {
        let body = '';
        res.on('data', (chunk) => {
          body += chunk;
        });
        res.on('end', () => {
          try {
            const response = JSON.parse(body);
            if (res.statusCode >= 200 && res.statusCode < 300) {
              resolve(response);
            } else {
              reject(new Error(`Razorpay API error: ${response.error?.description || body}`));
            }
          } catch (error) {
            reject(new Error(`Failed to parse Razorpay response: ${body}`));
          }
        });
      });

      req.on('error', (error) => {
        reject(new Error(`Razorpay request failed: ${error.message}`));
      });

      if (data && (method === 'POST' || method === 'PUT')) {
        req.write(JSON.stringify(data));
      }

      req.end();
    });
  }

  /**
   * Create a payout link for agent withdrawal
   * @param {Object} payoutData - Payout details
   * @param {number} payoutData.amount - Amount in rupees (₹)
   * @param {string} payoutData.account_holder_name - Account holder name
   * @param {string} payoutData.agent_email - Agent email
   * @returns {Promise<Object>} Payment link response
   */
  async createPayoutLink(payoutData) {
    try {
      const {
        amount,
        account_holder_name,
        agent_email,
        description = 'Agent Commission Withdrawal'
      } = payoutData;

      // Create a payment link for the agent to receive funds
      const paymentLink = await this.razorpay.paymentLink.create({
        amount: amount * 100, // Convert to paisa
        currency: 'INR',
        accept_partial: false,
        first_min_partial_amount: 0,
        expire_by: Math.floor(Date.now() / 1000) + (24 * 60 * 60), // 24 hours
        reference_id: `withdrawal_${Date.now()}`,
        description: description,
        customer: {
          name: account_holder_name,
          email: agent_email,
          contact: '+919876543210' // Dummy contact
        },
        notify: {
          sms: true,
          email: true
        },
        reminder_enable: true,
        notes: {
          type: 'agent_withdrawal',
          agent_name: account_holder_name
        },
        callback_url: `${config.FRONTEND_URL}/agent/withdrawals`,
        callback_method: 'get'
      });

      logger.info(`Created Razorpay payment link: ${paymentLink.id} for amount: ₹${amount}`);
      return paymentLink;

    } catch (error) {
      logger.error('Razorpay payment link creation failed:', error);
      throw new Error(`Payment link creation failed: ${error.message}`);
    }
  }

  /**
   * Create a payout to agent via UPI
   * @param {Object} payoutData - Payout details
   * @param {number} payoutData.amount - Amount in paisa (₹1 = 100 paisa)
   * @param {string} payoutData.upi_id - Agent's UPI ID
   * @param {string} payoutData.account_holder_name - Account holder name
   * @param {string} payoutData.agent_email - Agent email
   * @returns {Promise<Object>} Payout response
   */
  async createUpiPayout(payoutData) {
    try {
      const {
        amount,
        upi_id,
        account_holder_name,
        agent_email,
        narration = 'Agent Commission Withdrawal'
      } = payoutData;

      if (!this.isProduction) {
        // Development: Use mock payout
        return this.createMockUpiPayout(payoutData);
      } else {
        // Production: Use real Razorpay payout
        return this.createRealUpiPayout(payoutData);
      }
    } catch (error) {
      logger.error('UPI payout creation failed:', error);
      throw new Error(`UPI payout creation failed: ${error.message}`);
    }
  }

  /**
   * Create mock payout for development
   */
  async createMockUpiPayout(payoutData) {
    const { amount, upi_id, account_holder_name } = payoutData;
    const mockPayoutId = `pout_mock_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;

    logger.info(`[MOCK] UPI payout created: ${mockPayoutId} for amount: ₹${amount / 100} to UPI: ${upi_id}`);

    return {
      id: mockPayoutId,
      contact_id: `cont_mock_${Date.now()}`,
      fund_account_id: `fa_mock_${Date.now()}`,
      status: 'processed',
      amount: amount,
      currency: 'INR',
      upi_id: upi_id,
      created_at: Math.floor(Date.now() / 1000)
    };
  }

  /**
   * Create a payout to agent via Bank Transfer
   * @param {Object} payoutData - Payout details
   * @param {number} payoutData.amount - Amount in paisa (₹1 = 100 paisa)
   * @param {string} payoutData.account_number - Agent's bank account number
   * @param {string} payoutData.ifsc_code - IFSC code
   * @param {string} payoutData.account_holder_name - Account holder name
   * @param {string} payoutData.agent_email - Agent email
   * @returns {Promise<Object>} Payout response
   */
  async createBankPayout(payoutData) {
    try {
      const {
        amount,
        account_number,
        ifsc_code,
        account_holder_name,
        agent_email,
        narration = 'Agent Commission Withdrawal'
      } = payoutData;

      if (!this.isProduction) {
        // Development: Use mock payout
        return this.createMockBankPayout(payoutData);
      } else {
        // Production: Use real Razorpay payout
        return this.createRealBankPayout(payoutData);
      }
    } catch (error) {
      logger.error('Bank payout creation failed:', error);
      throw new Error(`Bank payout creation failed: ${error.message}`);
    }
  }

  /**
   * Create mock bank payout for development
   */
  async createMockBankPayout(payoutData) {
    const { amount, account_number, ifsc_code, account_holder_name } = payoutData;
    const mockPayoutId = `pout_bank_mock_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;

    logger.info(`[MOCK] Bank payout created: ${mockPayoutId} for amount: ₹${amount / 100} to A/C: ${account_number}`);

    return {
      id: mockPayoutId,
      contact_id: `cont_mock_${Date.now()}`,
      fund_account_id: `fa_mock_${Date.now()}`,
      status: 'processed',
      amount: amount,
      currency: 'INR',
      account_number: account_number,
      ifsc_code: ifsc_code,
      created_at: Math.floor(Date.now() / 1000)
    };
  }

  /**
   * Create real bank payout for production
   */
  async createRealBankPayout(payoutData) {
    const { amount, account_number, ifsc_code, account_holder_name, agent_email, narration } = payoutData;

    // Step 1: Create contact
    const contact = await this.makeRazorpayRequest('/contacts', 'POST', {
      name: account_holder_name,
      email: agent_email,
      contact: '9999999999',
      type: 'customer',
      reference_id: `agent_${Date.now()}`
    });

    // Step 2: Create fund account
    const fundAccount = await this.makeRazorpayRequest('/fund_accounts', 'POST', {
      contact_id: contact.id,
      account_type: 'bank_account',
      bank_account: {
        name: account_holder_name,
        account_number: account_number,
        ifsc: ifsc_code
      }
    });

    // Step 3: Create payout
    const payout = await this.makeRazorpayRequest('/payouts', 'POST', {
      account_number: config.RAZORPAY_ACCOUNT_NUMBER,
      fund_account_id: fundAccount.id,
      amount: amount,
      currency: 'INR',
      mode: 'IMPS',
      purpose: 'payout',
      queue_if_low_balance: true,
      reference_id: `withdrawal_${Date.now()}`,
      narration: narration
    });

    logger.info(`[PROD] Real bank payout created: ${payout.id} for amount: ₹${amount / 100} to A/C: ${account_number}`);

    return {
      id: payout.id,
      contact_id: contact.id,
      fund_account_id: fundAccount.id,
      status: payout.status,
      amount: amount,
      currency: 'INR',
      account_number: account_number,
      ifsc_code: ifsc_code,
      created_at: payout.created_at
    };
  }

  /**
   * Get payout details
   * @param {string} payoutId - Razorpay payout ID
   * @returns {Promise<Object>} Payout details
   */
  async getPayout(payoutId) {
    try {
      const payout = await this.razorpay.payouts.fetch(payoutId);
      return payout;
    } catch (error) {
      logger.error(`Failed to fetch payout ${payoutId}:`, error);
      throw new Error(`Failed to fetch payout: ${error.message}`);
    }
  }

  /**
   * Cancel a payout
   * @param {string} payoutId - Razorpay payout ID
   * @returns {Promise<Object>} Cancellation response
   */
  async cancelPayout(payoutId) {
    try {
      const response = await this.razorpay.payouts.cancel(payoutId);
      logger.info(`Cancelled payout: ${payoutId}`);
      return response;
    } catch (error) {
      logger.error(`Failed to cancel payout ${payoutId}:`, error);
      throw new Error(`Failed to cancel payout: ${error.message}`);
    }
  }

  /**
   * Verify webhook signature
   * @param {string} body - Raw request body
   * @param {string} signature - Razorpay signature
   * @param {string} secret - Webhook secret
   * @returns {boolean} Whether signature is valid
   */
  verifyWebhookSignature(body, signature, secret = config.RAZORPAY_WEBHOOK_SECRET) {
    try {
      const expectedSignature = crypto
        .createHmac('sha256', secret)
        .update(body)
        .digest('hex');

      return signature === expectedSignature;
    } catch (error) {
      logger.error('Webhook signature verification failed:', error);
      return false;
    }
  }
}

module.exports = new RazorpayService();