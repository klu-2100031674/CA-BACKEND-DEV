#!/usr/bin/env python3
import os
import sys
from openai import OpenAI

# Test the actual prompts used in report generation
def test_full_prompts():
    # Get API key from environment or command line
    api_key = os.getenv('PERPLEXITY_API_KEY') or (sys.argv[1] if len(sys.argv) > 1 else None)
    
    if not api_key:
        print("❌ No API key provided. Set PERPLEXITY_API_KEY environment variable or pass as argument.")
        return False
    
    # Initialize Perplexity client
    client = OpenAI(
        api_key=api_key,
        base_url="https://api.perplexity.ai/"
    )

    # Test prompts from the report generation
    test_prompts = [
        "Generate a comprehensive executive summary for a financial commission report covering Q4 2023. Include key metrics, trends, and strategic insights.",
        "Analyze the commission structure and provide detailed breakdown of calculations, rates, and payout distributions for the reporting period.",
        "Create a professional analysis section covering market performance, competitive positioning, and future outlook based on the commission data."
    ]

    print("🧪 Testing full report generation prompts with Perplexity API...")
    print(f"🔑 API Key: {api_key[:10]}...")
    print()

    for i, prompt in enumerate(test_prompts, 1):
        try:
            print(f"📝 Testing Prompt {i} (length: {len(prompt)} chars)")
            print(f"   Prompt: {prompt[:100]}...")

            import time
            time.sleep(2)  # Rate limiting delay

            response = client.chat.completions.create(
                model="sonar",
                messages=[{"role": "user", "content": prompt}]
            )

            content = response.choices[0].message.content
            print(f"   ✅ Success! Response length: {len(content)} chars")
            print(f"   📄 Preview: {content[:200]}...")
            print()

        except Exception as e:
            print(f"   ❌ Failed: {str(e)}")
            print(f"   Error type: {type(e).__name__}")
            print()
            return False

    print("🎉 All prompts tested successfully!")
    return True

if __name__ == "__main__":
    success = test_full_prompts()
    sys.exit(0 if success else 1)