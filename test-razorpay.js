require('dotenv').config();
const razorpayService = require('./src/services/razorpayService');

async function testRazorpayPaymentLink() {
  try {
    console.log('Testing Razorpay payment link creation...');

    const testData = {
      amount: 100, // ₹100
      account_holder_name: 'Test Agent',
      agent_email: 'test@example.com',
      description: 'Test Agent Commission Withdrawal'
    };

    console.log('Creating payment link with data:', testData);

    const paymentLink = await razorpayService.createPayoutLink(testData);

    console.log('Payment link created successfully:', paymentLink);

  } catch (error) {
    console.error('Payment link creation failed:', error.message);
    console.error('Full error:', error);
  }
}

// Run test
testRazorpayPaymentLink();