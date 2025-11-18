#!/usr/bin/env python3
"""
Test script for create_presentation_from_markdown

Tests the new Stage 3 API that converts Markdown to Google Slides.

Usage:
    python tests/slides/test_markdown_to_slides.py
"""

from mcp.client.streamable_http import streamablehttp_client
import asyncio
import json
import os
import sys
from pathlib import Path
from mcp import ClientSession

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

# Load environment variables
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    print("⚠️  python-dotenv not installed, using system environment variables only")

# Test configuration
ENDPOINT = "http://127.0.0.1:3333"


# Test Markdown content
SIMPLE_MARKDOWN = """
# My Product Launch

## Problem Statement
Current solutions are:
- Expensive
- Complex
- Slow

## Our Solution
We offer:
- Affordable pricing
- Simple interface
- Fast performance

## Conclusion
Thank you for your attention!
"""

MARKDOWN_WITH_IMAGES = """
# AI Technology Overview

## What is AI?
Artificial Intelligence is transforming the world.

![AI Diagram](https://images.netcomlearning.com/cms/banners/what-is-ai-blog-banner.jpg)

Key areas:
- Machine Learning
- Deep Learning
- Neural Networks

## Use Cases
AI is used in:
1. Healthcare
2. Finance
3. Transportation

## Future Outlook
The future is bright!
"""

MARKDOWN_WITH_CODE = """
# API Documentation

## Installation
Install the package:

```bash
pip install my-package
```

## Usage Example
Here's a simple example:

```python
def hello():
    print("Hello World")
```

## Configuration
Set these environment variables:
- API_KEY
- SECRET_TOKEN
"""

COMPLEX_MARKDOWN = """
# Quarterly Report Q4 2024

## Executive Summary
This quarter showed strong growth across all metrics.

### Key Highlights
- Revenue increased 25%
- Customer satisfaction at 95%
- New product launch successful

## Financial Performance
![Revenue Chart](https://www.shutterstock.com/image-vector/galati-romania-april-29-2023-260nw-2295394661.jpg)

Revenue breakdown:
1. Product Sales: $5M
2. Services: $3M
3. Subscriptions: $2M

### Expenses
Operating costs were well-managed:
- Personnel: 40%
- Infrastructure: 30%
- Marketing: 20%
- R&D: 10%

## Customer Metrics
Customer growth exceeded expectations.

Acquisition channels:
- Organic search: 45%
- Referrals: 30%
- Paid ads: 25%

## Future Plans
Looking ahead to Q1 2025:
- Launch mobile app
- Expand to 5 new markets
- Hire 20 new team members
"""


async def test_markdown_to_slides():
    """Test the create_presentation_from_markdown API"""
    print("=" * 80)
    print("🎯 Testing create_presentation_from_markdown (Stage 3)")
    print("=" * 80)

    # OAuth headers
    headers = {
        "GOOGLE_OAUTH_REFRESH_TOKEN": os.getenv("TEST_GOOGLE_OAUTH_REFRESH_TOKEN"),
        "GOOGLE_OAUTH_CLIENT_ID": os.getenv("TEST_GOOGLE_OAUTH_CLIENT_ID"),
        "GOOGLE_OAUTH_CLIENT_SECRET": os.getenv("TEST_GOOGLE_OAUTH_CLIENT_SECRET"),
    }

    if not all(headers.values()):
        print("❌ Missing required environment variables")
        return False

    async with streamablehttp_client(url=f"{ENDPOINT}/mcp", headers=headers) as (read, write, _):
        async with ClientSession(read, write) as session:
            await session.initialize()

            # Test 1: Simple markdown (3 slides)
            print("\n📝 Test 1: Simple markdown with H1 and H2s")
            result = await session.call_tool("create_presentation_from_markdown", {
                "markdown_content": SIMPLE_MARKDOWN
            })

            if result.isError:
                print(f"❌ Failed: {result.content[0].text if result.content else 'Unknown error'}")
                return False

            response = json.loads(result.content[0].text)
            print(f"✅ Created presentation:")
            print(f"   Title: {response['presentation_title']}")
            print(f"   URL: {response['presentation_url']}")
            print(f"   Slides created: {response['slides_created']}")
            print(f"   Total elements: {response['total_elements']}")
            if response.get('warnings'):
                print(f"   Warnings: {len(response['warnings'])}")
                for warning in response['warnings']:
                    print(f"     - {warning}")

            assert response['slides_created'] == 3, f"Expected 3 slides, got {response['slides_created']}"

            # Test 2: Markdown with images
            print("\n📝 Test 2: Markdown with images")
            result = await session.call_tool("create_presentation_from_markdown", {
                "markdown_content": MARKDOWN_WITH_IMAGES
            })

            if result.isError:
                print(f"❌ Failed: {result.content[0].text if result.content else 'Unknown error'}")
                return False

            response = json.loads(result.content[0].text)
            print(f"✅ Created presentation with images:")
            print(f"   Title: {response['presentation_title']}")
            print(f"   URL: {response['presentation_url']}")
            print(f"   Slides created: {response['slides_created']}")
            print(f"   Total elements: {response['total_elements']}")
            if response.get('warnings'):
                print(f"   Warnings: {len(response['warnings'])}")

            # Test 3: Markdown with code blocks
            print("\n📝 Test 3: Markdown with code blocks")
            result = await session.call_tool("create_presentation_from_markdown", {
                "markdown_content": MARKDOWN_WITH_CODE
            })

            if result.isError:
                print(f"❌ Failed: {result.content[0].text if result.content else 'Unknown error'}")
                return False

            response = json.loads(result.content[0].text)
            print(f"✅ Created presentation with code blocks:")
            print(f"   Title: {response['presentation_title']}")
            print(f"   URL: {response['presentation_url']}")
            print(f"   Slides created: {response['slides_created']}")
            print(f"   Total elements: {response['total_elements']}")

            # Test 4: Complex markdown with H3, lists, images
            print("\n📝 Test 4: Complex markdown with multiple features")
            result = await session.call_tool("create_presentation_from_markdown", {
                "markdown_content": COMPLEX_MARKDOWN
            })

            if result.isError:
                print(f"❌ Failed: {result.content[0].text if result.content else 'Unknown error'}")
                return False

            response = json.loads(result.content[0].text)
            print(f"✅ Created complex presentation:")
            print(f"   Title: {response['presentation_title']}")
            print(f"   URL: {response['presentation_url']}")
            print(f"   Slides created: {response['slides_created']}")
            print(f"   Total elements: {response['total_elements']}")

            # Test 5: Override title
            print("\n📝 Test 5: Override presentation title")
            result = await session.call_tool("create_presentation_from_markdown", {
                "markdown_content": SIMPLE_MARKDOWN,
                "presentation_title": "Custom Title Override"
            })

            if result.isError:
                print(f"❌ Failed: {result.content[0].text if result.content else 'Unknown error'}")
                return False

            response = json.loads(result.content[0].text)
            print(f"✅ Created presentation with custom title:")
            print(f"   Title: {response['presentation_title']}")
            print(f"   URL: {response['presentation_url']}")
            assert response['presentation_title'] == "Custom Title Override"

            # Test 6: Error case - no H1
            print("\n📝 Test 6: Error case - no H1 heading")
            no_h1_markdown = """
## First Slide
Content without H1

## Second Slide
More content
"""
            result = await session.call_tool("create_presentation_from_markdown", {
                "markdown_content": no_h1_markdown
            })

            if result.isError:
                print(f"✅ Correctly returned error (MCP error): {result.content[0].text if result.content else 'Unknown error'}")
            else:
                response = json.loads(result.content[0].text)
                if not response.get('success'):
                    print(f"✅ Correctly returned error: {response.get('error')}")
                else:
                    print(f"⚠️  Should have failed but succeeded")
                    return False

            # Test 7: Error case - no H2
            print("\n📝 Test 7: Error case - only H1, no H2 slides")
            only_h1_markdown = """
# Just a Title
Some content but no slides
"""
            result = await session.call_tool("create_presentation_from_markdown", {
                "markdown_content": only_h1_markdown
            })

            if result.isError:
                print(f"✅ Correctly returned error (MCP error): {result.content[0].text if result.content else 'Unknown error'}")
            else:
                try:
                    response = json.loads(result.content[0].text)
                    if not response.get('success'):
                        print(f"✅ Correctly returned error: {response.get('error')}")
                    else:
                        print(f"⚠️  Should have failed but succeeded")
                        return False
                except json.JSONDecodeError as e:
                    print(f"❌ JSON decode error: {e}")
                    print(f"   Response was: {result.content[0].text if result.content else 'empty'}")
                    return False

            print("\n" + "=" * 80)
            print("✅ ALL TESTS PASSED!")
            print("=" * 80)
            print(f"\n📊 Test Summary:")
            print(f"   ✅ Simple markdown")
            print(f"   ✅ Markdown with images")
            print(f"   ✅ Markdown with code blocks")
            print(f"   ✅ Complex markdown")
            print(f"   ✅ Custom title override")
            print(f"   ✅ Error handling (no H1)")
            print(f"   ✅ Error handling (no H2)")

            return True


if __name__ == "__main__":
    success = asyncio.run(test_markdown_to_slides())
    sys.exit(0 if success else 1)
