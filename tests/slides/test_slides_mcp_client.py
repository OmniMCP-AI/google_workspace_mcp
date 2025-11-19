#!/usr/bin/env python3
"""
Google Slides MCP Client Integration Test

This test uses the MCP client to call the Slides server via HTTP,
avoiding direct Google API authentication issues.

Usage:
    # Test against local MCP server
    python tests/test_slides_mcp_client.py --env=local --test=add_text

    # Test against production server
    python tests/test_slides_mcp_client.py --env=prod --test=add_text

Environment Variables Required:
- TEST_GOOGLE_OAUTH_REFRESH_TOKEN
- TEST_GOOGLE_OAUTH_CLIENT_ID
- TEST_GOOGLE_OAUTH_CLIENT_SECRET
"""

from mcp.client.streamable_http import streamablehttp_client
import asyncio
import json
import argparse
import os
import sys
from pathlib import Path
from mcp import ClientSession
from datetime import datetime

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

# Test configuration
TEST_USER_ID = "68501372a3569b6897673a48"
TEST_PRESENTATION_URL = "https://docs.google.com/presentation/d/1RjD6lUqincqSZvwilTosA3IxrY7DSmph7CBECvWLQYg/edit?slide=id.p#slide=id.p"


async def test_add_text(url, headers):
    """Test adding text boxes to a slide using batch_update_presentation"""
    print("=" * 80)
    print("🎯 Test: Adding Text Boxes via MCP Client")
    print("=" * 80)

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    async with streamablehttp_client(url=url, headers=headers) as (read, write, _):
        async with ClientSession(read, write) as session:
            await session.initialize()

            # Test 0: List available tools
            print(f"\n🛠️  Step 0: Listing available MCP tools")
            tools = await session.list_tools()
            print(f"✅ Found {len(tools.tools)} available tools:")
            for i, tool in enumerate(tools.tools, 1):
                if 'slide' in tool.name.lower() or 'presentation' in tool.name.lower():
                    print(f"   {i:2d}. {tool.name}: {tool.description[:80]}...")
            print()

            # Test 1: Get presentation metadata first
            print(f"\n📘 Step 1: Getting presentation metadata")
            print(f"   URL: {TEST_PRESENTATION_URL}")

            # Note: We'll use batch_update_presentation directly
            # First, let's create a simple text box

            # Create text box with title
            text_box_id = f"textbox_{timestamp.replace(' ', '_').replace(':', '_')}"

            requests = [
                # Create a text box
                {
                    "createShape": {
                        "objectId": text_box_id,
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "size": {
                                "height": {"magnitude": 100, "unit": "PT"},
                                "width": {"magnitude": 400, "unit": "PT"}
                            },
                            "transform": {
                                "scaleX": 1,
                                "scaleY": 1,
                                "translateX": 50,
                                "translateY": 50,
                                "unit": "PT"
                            }
                        }
                    }
                },
                # Insert text into the text box
                {
                    "insertText": {
                        "objectId": text_box_id,
                        "text": f"Test Text Box via MCP Client - {timestamp}",
                        "insertionIndex": 0
                    }
                }
            ]

            print(f"\n📝 Step 2: Creating text box via batch_update_presentation")
            print(f"   Text box ID: {text_box_id}")
            print(f"   Operations: {len(requests)}")

            # Call the batch_update_presentation tool
            result = await session.call_tool("batch_update_presentation", {
                "presentation_url": TEST_PRESENTATION_URL,
                "requests": requests
            })

            print(f"\n📊 Result: {result}")

            # Verify the result
            if result.isError:
                print(f"❌ FAIL: Error occurred")
                if result.content and result.content[0].text:
                    print(f"   Error: {result.content[0].text}")
                return False
            elif result.content and result.content[0].text:
                content = json.loads(result.content[0].text)

                if content.get('success'):
                    print(f"✅ SUCCESS: Text box created!")
                    print(f"   Presentation ID: {content.get('presentation_id', 'N/A')}")
                    print(f"   Replies: {len(content.get('replies', []))}")

                    # Check replies for created shape
                    replies = content.get('replies', [])
                    for reply in replies:
                        if 'createShape' in reply:
                            created_id = reply['createShape'].get('objectId')
                            print(f"   Created shape ID: {created_id}")

                    return True
                else:
                    print(f"❌ FAIL: {content.get('message', 'Unknown error')}")
                    return False
            else:
                print(f"❌ FAIL: No content in response")
                return False


async def test_create_slides(url, headers):
    """Test creating new slides using batch_update_presentation"""
    print("=" * 80)
    print("🎯 Test: Creating Multiple Slides via MCP Client")
    print("=" * 80)

    async with streamablehttp_client(url=url, headers=headers) as (read, write, _):
        async with ClientSession(read, write) as session:
            await session.initialize()

            print(f"\n📝 Creating 3 new slides...")

            # Create requests to add 3 new slides
            requests = [
                {
                    "createSlide": {
                        "slideLayoutReference": {
                            "predefinedLayout": "TITLE_AND_BODY"
                        }
                    }
                }
                for i in range(3)
            ]

            # Call the batch_update_presentation tool
            result = await session.call_tool("batch_update_presentation", {
                "presentation_url": TEST_PRESENTATION_URL,
                "requests": requests
            })

            # Verify the result
            if result.isError:
                print(f"❌ FAIL: Error occurred")
                if result.content and result.content[0].text:
                    print(f"   Error: {result.content[0].text}")
                return False
            elif result.content and result.content[0].text:
                content = json.loads(result.content[0].text)

                if content.get('success'):
                    replies = content.get('replies', [])
                    print(f"✅ SUCCESS: Created {len(replies)} slides!")

                    # List created slide IDs
                    for i, reply in enumerate(replies, 1):
                        if 'createSlide' in reply:
                            slide_id = reply['createSlide'].get('objectId', 'Unknown')
                            print(f"   Slide {i}: {slide_id}")

                    return True
                else:
                    print(f"❌ FAIL: {content.get('message', 'Unknown error')}")
                    return False


async def run_single_test(test_name, url, headers):
    """Run a single test by name"""
    test_functions = {
        'add_text': test_add_text,
        'create_slides': test_create_slides,
    }

    if test_name not in test_functions:
        print(f"❌ Unknown test: {test_name}")
        print(f"Available tests: {', '.join(test_functions.keys())}")
        return

    print(f"\n🎯 Running test: {test_name}")

    try:
        result = await test_functions[test_name](url, headers)

        print("\n" + "=" * 80)
        if result:
            print("✅ TEST PASSED")
        else:
            print("❌ TEST FAILED")
        print("=" * 80)

        return result
    except Exception as e:
        print(f"\n❌ Test '{test_name}' failed with error: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    # Load environment variables
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except ImportError:
        print("⚠️  python-dotenv not installed, using system environment variables only")

    # Parse command line arguments
    parser = argparse.ArgumentParser(description="Test Google Slides MCP Integration")
    parser.add_argument(
        "--env",
        choices=["local", "prod"],
        default="local",
        help="Environment: local (127.0.0.1:8321) or prod"
    )
    parser.add_argument(
        "--test",
        choices=["add_text", "create_slides"],
        default="add_text",
        help="Which test to run"
    )
    args = parser.parse_args()

    # Set endpoint based on environment
    if args.env == "prod":
        # Replace with your actual production endpoint
        endpoint = "https://your-slides-mcp-server.com"
    else:
        endpoint = "http://127.0.0.1:8321"

    print(f"🔗 Using {args.env} environment: {endpoint}")
    print(f"📋 User ID: {TEST_USER_ID}")

    # OAuth headers for authentication
    test_headers = {
        "GOOGLE_OAUTH_REFRESH_TOKEN": os.getenv("TEST_GOOGLE_OAUTH_REFRESH_TOKEN"),
        "GOOGLE_OAUTH_CLIENT_ID": os.getenv("TEST_GOOGLE_OAUTH_CLIENT_ID"),
        "GOOGLE_OAUTH_CLIENT_SECRET": os.getenv("TEST_GOOGLE_OAUTH_CLIENT_SECRET"),
    }

    # Verify environment variables are set
    if not all(test_headers.values()):
        print("❌ Missing required environment variables:")
        print("   - TEST_GOOGLE_OAUTH_REFRESH_TOKEN")
        print("   - TEST_GOOGLE_OAUTH_CLIENT_ID")
        print("   - TEST_GOOGLE_OAUTH_CLIENT_SECRET")
        print("\nPlease set these in your .env file")
        sys.exit(1)

    # Run the test
    asyncio.run(run_single_test(args.test, url=f"{endpoint}/mcp", headers=test_headers))
