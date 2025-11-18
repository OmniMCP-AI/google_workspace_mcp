#!/usr/bin/env python3
"""
Test script for add_page_with_content convenience wrapper

Tests the new Stage 2 API that combines slide creation with content addition.

Usage:
    python tests/slides/test_add_page_with_content.py
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
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

# Load environment variables
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    print("⚠️  python-dotenv not installed, using system environment variables only")

# Test configuration
ENDPOINT = "http://127.0.0.1:3333"


async def test_add_page_with_content():
    """Test the new add_page_with_content convenience wrapper"""
    print("=" * 80)
    print("🎯 Testing add_page_with_content (Stage 2 Convenience Wrapper)")
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

            # Test 0: Create presentation
            print("\n📝 Test 0: Creating presentation")
            result = await session.call_tool("create_presentation", {
                "title": f"Stage 2 Test - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            })

            if result.isError:
                print(f"❌ Failed: {result.content[0].text if result.content else 'Unknown error'}")
                return False

            response = json.loads(result.content[0].text)
            presentation_url = response['presentation_url']
            print(f"✅ Created presentation: {presentation_url}")

            # Test 1: Full slide (title + body + image)
            print("\n📝 Test 1: Creating full slide (title + body + image)")
            result = await session.call_tool("add_page_with_content", {
                "presentation_url": presentation_url,
                "title": "AI Hot News - Full Slide Test",
                "body_text": "This slide has:\n• Title at top\n• Body text here\n• Image on the right",
                "image_url": "https://static-cdn.toi-media.com/www/uploads/2014/07/gal-gadot.jpg"
            })

            if result.isError:
                print(f"❌ Failed: {result.content[0].text if result.content else 'Unknown error'}")
                return False

            response = json.loads(result.content[0].text)
            print(f"✅ Full slide created:")
            print(f"   Slide ID: {response['slide_id']}")
            print(f"   Elements added: {len(response['elements_added'])}")
            for elem in response['elements_added']:
                print(f"     - {elem['type']}: {elem['object_id']}")

            # Test 2: Title only
            print("\n📝 Test 2: Creating slide with title only")
            result = await session.call_tool("add_page_with_content", {
                "presentation_url": presentation_url,
                "title": "Section Header - Title Only"
            })

            if result.isError:
                print(f"❌ Failed: {result.content[0].text if result.content else 'Unknown error'}")
                return False

            response = json.loads(result.content[0].text)
            print(f"✅ Title-only slide created:")
            print(f"   Slide ID: {response['slide_id']}")
            print(f"   Elements added: {len(response['elements_added'])}")
            assert len(response['elements_added']) == 1, "Should have exactly 1 element (title)"
            assert response['elements_added'][0]['type'] == 'title', "Element should be title"

            # Test 3: Body text only
            print("\n📝 Test 3: Creating slide with body text only")
            result = await session.call_tool("add_page_with_content", {
                "presentation_url": presentation_url,
                "body_text": "This slide has only body text.\n\nNo title, no image.\n\nJust content."
            })

            if result.isError:
                print(f"❌ Failed: {result.content[0].text if result.content else 'Unknown error'}")
                return False

            response = json.loads(result.content[0].text)
            print(f"✅ Body-only slide created:")
            print(f"   Slide ID: {response['slide_id']}")
            print(f"   Elements added: {len(response['elements_added'])}")
            assert len(response['elements_added']) == 1, "Should have exactly 1 element (body)"
            assert response['elements_added'][0]['type'] == 'body_text', "Element should be body_text"

            # Test 4: Image only
            print("\n📝 Test 4: Creating slide with image only")
            result = await session.call_tool("add_page_with_content", {
                "presentation_url": presentation_url,
                "image_url": "https://static-cdn.toi-media.com/www/uploads/2014/07/gal-gadot.jpg",
                "image_width": 500,
                "image_height": 300,
                "image_x": 150,
                "image_y": 100
            })

            if result.isError:
                print(f"❌ Failed: {result.content[0].text if result.content else 'Unknown error'}")
                return False

            response = json.loads(result.content[0].text)
            print(f"✅ Image-only slide created:")
            print(f"   Slide ID: {response['slide_id']}")
            print(f"   Elements added: {len(response['elements_added'])}")
            assert len(response['elements_added']) == 1, "Should have exactly 1 element (image)"
            assert response['elements_added'][0]['type'] == 'image', "Element should be image"

            # Test 5: Custom positioning
            print("\n📝 Test 5: Creating slide with custom positioning")
            result = await session.call_tool("add_page_with_content", {
                "presentation_url": presentation_url,
                "title": "Custom Layout",
                "title_x": 100,
                "title_y": 50,
                "title_width": 500,
                "title_height": 60,
                "body_text": "Custom positioned content",
                "body_x": 100,
                "body_y": 150,
                "body_width": 500,
                "body_height": 150
            })

            if result.isError:
                print(f"❌ Failed: {result.content[0].text if result.content else 'Unknown error'}")
                return False

            response = json.loads(result.content[0].text)
            print(f"✅ Custom layout slide created:")
            print(f"   Slide ID: {response['slide_id']}")
            print(f"   Elements added: {len(response['elements_added'])}")

            # Test 6: Title + Image (no body)
            print("\n📝 Test 6: Creating slide with title + image (no body)")
            result = await session.call_tool("add_page_with_content", {
                "presentation_url": presentation_url,
                "title": "Image Gallery Item",
                "image_url": "https://static-cdn.toi-media.com/www/uploads/2014/07/gal-gadot.jpg",
                "image_y": 150
            })

            if result.isError:
                print(f"❌ Failed: {result.content[0].text if result.content else 'Unknown error'}")
                return False

            response = json.loads(result.content[0].text)
            print(f"✅ Title + image slide created:")
            print(f"   Slide ID: {response['slide_id']}")
            print(f"   Elements added: {len(response['elements_added'])}")
            assert len(response['elements_added']) == 2, "Should have exactly 2 elements"

            # Test 7: Empty slide (no content)
            print("\n📝 Test 7: Creating empty slide (no content)")
            result = await session.call_tool("add_page_with_content", {
                "presentation_url": presentation_url,
                "layout": "BLANK"
            })

            if result.isError:
                print(f"❌ Failed: {result.content[0].text if result.content else 'Unknown error'}")
                return False

            response = json.loads(result.content[0].text)
            print(f"✅ Empty slide created:")
            print(f"   Slide ID: {response['slide_id']}")
            print(f"   Elements added: {len(response['elements_added'])}")
            assert len(response['elements_added']) == 0, "Should have no elements"

            # Test 8: Using different layout
            print("\n📝 Test 8: Creating slide with TITLE_AND_BODY layout")
            result = await session.call_tool("add_page_with_content", {
                "presentation_url": presentation_url,
                "layout": "TITLE_AND_BODY",
                "title": "Using Predefined Layout",
                "body_text": "This uses TITLE_AND_BODY layout instead of BLANK"
            })

            if result.isError:
                print(f"❌ Failed: {result.content[0].text if result.content else 'Unknown error'}")
                return False

            response = json.loads(result.content[0].text)
            print(f"✅ TITLE_AND_BODY layout slide created:")
            print(f"   Slide ID: {response['slide_id']}")
            print(f"   Layout: {response['layout']}")
            print(f"   Elements added: {len(response['elements_added'])}")

            print("\n" + "=" * 80)
            print("✅ ALL TESTS PASSED!")
            print("=" * 80)
            print(f"\n📊 View your presentation:")
            print(f"   {presentation_url}")
            print(f"\n✅ Summary:")
            print(f"   - Test 1: Full slide (3 elements)")
            print(f"   - Test 2: Title only (1 element)")
            print(f"   - Test 3: Body only (1 element)")
            print(f"   - Test 4: Image only (1 element)")
            print(f"   - Test 5: Custom positioning (2 elements)")
            print(f"   - Test 6: Title + image (2 elements)")
            print(f"   - Test 7: Empty slide (0 elements)")
            print(f"   - Test 8: Different layout (2 elements)")

            return True


if __name__ == "__main__":
    success = asyncio.run(test_add_page_with_content())
    sys.exit(0 if success else 1)
