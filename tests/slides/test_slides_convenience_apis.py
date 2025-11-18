#!/usr/bin/env python3
"""
Test script for Google Slides convenience APIs

Tests the new high-level APIs: add_slide, add_title, add_body_text, add_body_image

Usage:
    python tests/test_slides_convenience_apis.py
"""

from mcp.client.streamable_http import streamablehttp_client
import asyncio
import json
import os
import sys
from pathlib import Path
from mcp import ClientSession
from datetime import datetime

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

# Load environment variables
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    print("⚠️  python-dotenv not installed, using system environment variables only")

# Test configuration
TEST_USER_ID = "68501372a3569b6897673a48"
ENDPOINT = "http://127.0.0.1:3333"


async def test_convenience_apis():
    """Test the new convenience APIs"""
    print("=" * 80)
    print("🎯 Testing Google Slides Convenience APIs")
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

            # Test 1: Create presentation
            print("\n📝 Test 1: Creating presentation with structured response")
            result = await session.call_tool("create_presentation", {
                "title": f"Convenience API Test - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            })

            if result.isError:
                print(f"❌ Failed: {result.content[0].text if result.content else 'Unknown error'}")
                return False

            response = json.loads(result.content[0].text)
            print(f"✅ Created presentation:")
            print(f"   ID: {response['presentation_id']}")
            print(f"   URL: {response['presentation_url']}")

            presentation_url = response['presentation_url']

            # Test 2: Add slides
            print("\n📝 Test 2: Adding 3 slides")
            slide_ids = []
            for i in range(3):
                result = await session.call_tool("add_slide", {
                    "presentation_url": presentation_url,
                    "layout": "BLANK"
                })

                if result.isError:
                    print(f"❌ Failed to add slide {i+1}")
                    return False

                response = json.loads(result.content[0].text)
                slide_ids.append(response['slide_id'])
                print(f"   ✅ Slide {i+1} created: {response['slide_id']}")

            # Test 3: Add title to first slide
            print("\n📝 Test 3: Adding title to first slide")
            result = await session.call_tool("add_title", {
                "presentation_url": presentation_url,
                "text": "Welcome to Convenience APIs",
                "page_id": "first"
            })

            if result.isError:
                print(f"❌ Failed: {result.content[0].text if result.content else 'Unknown error'}")
                return False

            response = json.loads(result.content[0].text)
            print(f"   ✅ Title added: {response['title_object_id']}")

            # Test 4: Add body text to first slide
            print("\n📝 Test 4: Adding body text to first slide")
            result = await session.call_tool("add_body_text", {
                "presentation_url": presentation_url,
                "text": "This presentation demonstrates the new convenience APIs:\n\n• add_slide\n• add_title\n• add_body_text\n• add_body_image",
                "page_id": "first"
            })

            if result.isError:
                print(f"❌ Failed: {result.content[0].text if result.content else 'Unknown error'}")
                return False

            response = json.loads(result.content[0].text)
            print(f"   ✅ Body text added: {response['body_object_id']}")

            # Test 5: Add image to last slide
            print("\n📝 Test 5: Adding image to last slide")
            result = await session.call_tool("add_body_image", {
                "presentation_url": presentation_url,
                "image_url": "https://static-cdn.toi-media.com/www/uploads/2014/07/gal-gadot.jpg",
                "page_id": "last"
            })

            if result.isError:
                print(f"❌ Failed: {result.content[0].text if result.content else 'Unknown error'}")
                return False

            response = json.loads(result.content[0].text)
            print(f"   ✅ Image added: {response['image_object_id']}")

            # Test 6: Add title and body to middle slide
            print("\n📝 Test 6: Adding title and body to second slide")
            result = await session.call_tool("add_title", {
                "presentation_url": presentation_url,
                "text": "Slide 2: Testing Positions",
                "page_id": slide_ids[1]  # Use specific slide ID
            })

            if result.isError:
                print(f"❌ Failed to add title")
                return False

            result = await session.call_tool("add_body_text", {
                "presentation_url": presentation_url,
                "text": "This slide uses a specific slide ID instead of 'first' or 'last'",
                "page_id": slide_ids[1]
            })

            if result.isError:
                print(f"❌ Failed to add body")
                return False

            print(f"   ✅ Added title and body to specific slide: {slide_ids[1]}")

            print("\n" + "=" * 80)
            print("✅ ALL TESTS PASSED!")
            print("=" * 80)
            print(f"\n📊 View your presentation:")
            print(f"   {presentation_url}")

            return True


if __name__ == "__main__":
    success = asyncio.run(test_convenience_apis())
    sys.exit(0 if success else 1)
