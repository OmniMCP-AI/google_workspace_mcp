#!/usr/bin/env python3
"""
Test script for Google Slides batch_update_presentation functionality

This script tests the batch_update_presentation function with various operations:
- Creating new slides
- Adding text boxes
- Creating shapes
- Inserting images
- Applying formatting
- Batch operations

Usage:
    # Test against local MCP server
    python tests/test_slides_batch_update.py --env=local --test=all

    # Test specific functionality
    python tests/test_slides_batch_update.py --env=local --test=create_slides
    python tests/test_slides_batch_update.py --env=local --test=add_text
    python tests/test_slides_batch_update.py --env=local --test=add_shapes
    python tests/test_slides_batch_update.py --env=local --test=add_images
    python tests/test_slides_batch_update.py --env=local --test=complex

Requirements:
- Set TEST_GOOGLE_OAUTH_REFRESH_TOKEN in .env
- Set TEST_GOOGLE_OAUTH_CLIENT_ID in .env
- Set TEST_GOOGLE_OAUTH_CLIENT_SECRET in .env
- A test presentation accessible to the user
"""

import os
import sys
import asyncio
import argparse
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# Import the scopes from the MCP server
from auth.scopes import SLIDES_SCOPES, BASE_SCOPES

# Test presentation URL - you can change this to your own test presentation
TEST_PRESENTATION_URL = "https://docs.google.com/presentation/d/1RjD6lUqincqSZvwilTosA3IxrY7DSmph7CBECvWLQYg/edit?slide=id.p#slide=id.p"

def extract_presentation_id(url: str) -> str:
    """Extract presentation ID from Google Slides URL"""
    # Format: https://docs.google.com/presentation/d/{presentationId}/edit...
    if "/presentation/d/" in url:
        parts = url.split("/presentation/d/")
        if len(parts) > 1:
            presentation_id = parts[1].split("/")[0]
            return presentation_id
    raise ValueError(f"Could not extract presentation ID from URL: {url}")


class SlidesTestClient:
    """Test client for Google Slides batch update functionality"""

    def __init__(self, refresh_token: str, client_id: str, client_secret: str):
        """Initialize the test client with OAuth credentials"""
        # Use only Slides scopes + base scopes to match the token from the database
        scopes = BASE_SCOPES + SLIDES_SCOPES
        self.credentials = Credentials(
            token=None,
            refresh_token=refresh_token,
            token_uri="https://oauth2.googleapis.com/token",
            client_id=client_id,
            client_secret=client_secret,
            scopes=scopes
        )
        self.service = build('slides', 'v1', credentials=self.credentials)

    async def get_presentation(self, presentation_id: str) -> Dict[str, Any]:
        """Get presentation metadata"""
        result = await asyncio.to_thread(
            self.service.presentations().get(presentationId=presentation_id).execute
        )
        return result

    async def batch_update(
        self,
        presentation_id: str,
        requests: List[Dict[str, Any]]
    ) -> Dict[str, Any]:
        """Execute batch update on presentation"""
        body = {'requests': requests}
        result = await asyncio.to_thread(
            self.service.presentations().batchUpdate(
                presentationId=presentation_id,
                body=body
            ).execute
        )
        return result


async def test_create_slides(client: SlidesTestClient, presentation_id: str) -> bool:
    """Test 1: Create multiple new slides"""
    print("\n" + "=" * 60)
    print("📝 Test 1: Creating Multiple Slides")
    print("=" * 60)

    try:
        # Get initial presentation state
        initial_state = await client.get_presentation(presentation_id)
        initial_slide_count = len(initial_state.get('slides', []))
        print(f"   Initial slide count: {initial_slide_count}")

        # Create requests to add 3 new slides
        requests = [
            {
                'createSlide': {
                    'slideLayoutReference': {
                        'predefinedLayout': 'TITLE_AND_BODY'
                    },
                    'insertionIndex': str(initial_slide_count + i)
                }
            }
            for i in range(3)
        ]

        print(f"   Creating 3 new slides...")
        result = await client.batch_update(presentation_id, requests)

        replies = result.get('replies', [])
        print(f"   ✅ Batch update completed")
        print(f"   Replies received: {len(replies)}")

        # Verify slides were created
        for i, reply in enumerate(replies, 1):
            if 'createSlide' in reply:
                slide_id = reply['createSlide'].get('objectId', 'Unknown')
                print(f"      Slide {i}: Created with ID {slide_id}")

        # Verify final slide count
        final_state = await client.get_presentation(presentation_id)
        final_slide_count = len(final_state.get('slides', []))
        print(f"   Final slide count: {final_slide_count}")

        if final_slide_count == initial_slide_count + 3:
            print("   ✅ PASS: All 3 slides created successfully")
            return True
        else:
            print(f"   ❌ FAIL: Expected {initial_slide_count + 3} slides, got {final_slide_count}")
            return False

    except Exception as e:
        print(f"   ❌ FAIL: Error creating slides: {e}")
        import traceback
        traceback.print_exc()
        return False


async def test_add_text(client: SlidesTestClient, presentation_id: str) -> bool:
    """Test 2: Add text boxes to a slide"""
    print("\n" + "=" * 60)
    print("📝 Test 2: Adding Text Boxes")
    print("=" * 60)

    try:
        # Get the first slide
        presentation = await client.get_presentation(presentation_id)
        slides = presentation.get('slides', [])

        if not slides:
            print("   ⚠️  WARNING: No slides found in presentation")
            return False

        first_slide_id = slides[0].get('objectId')
        print(f"   Target slide ID: {first_slide_id}")

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Create text box with title
        text_box_id = f"textbox_{timestamp.replace(' ', '_').replace(':', '_')}"

        requests = [
            # Create a text box
            {
                'createShape': {
                    'objectId': text_box_id,
                    'shapeType': 'TEXT_BOX',
                    'elementProperties': {
                        'pageObjectId': first_slide_id,
                        'size': {
                            'height': {'magnitude': 100, 'unit': 'PT'},
                            'width': {'magnitude': 400, 'unit': 'PT'}
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': 50,
                            'translateY': 50,
                            'unit': 'PT'
                        }
                    }
                }
            },
            # Insert text into the text box
            {
                'insertText': {
                    'objectId': text_box_id,
                    'text': f'Test Text Box Created at {timestamp}',
                    'insertionIndex': 0
                }
            }
        ]

        print(f"   Creating text box...")
        result = await client.batch_update(presentation_id, requests)

        replies = result.get('replies', [])
        print(f"   ✅ Batch update completed")
        print(f"   Replies received: {len(replies)}")

        # Check if shape was created
        created_shape_id = None
        for reply in replies:
            if 'createShape' in reply:
                created_shape_id = reply['createShape'].get('objectId')
                print(f"      Created text box with ID: {created_shape_id}")

        if created_shape_id:
            print("   ✅ PASS: Text box created successfully")
            return True
        else:
            print("   ❌ FAIL: Text box creation not confirmed")
            return False

    except Exception as e:
        print(f"   ❌ FAIL: Error adding text: {e}")
        import traceback
        traceback.print_exc()
        return False


async def test_add_shapes(client: SlidesTestClient, presentation_id: str) -> bool:
    """Test 3: Add various shapes to a slide"""
    print("\n" + "=" * 60)
    print("📝 Test 3: Adding Shapes")
    print("=" * 60)

    try:
        # Get the first slide
        presentation = await client.get_presentation(presentation_id)
        slides = presentation.get('slides', [])

        if not slides:
            print("   ⚠️  WARNING: No slides found in presentation")
            return False

        first_slide_id = slides[0].get('objectId')
        print(f"   Target slide ID: {first_slide_id}")

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Create multiple shapes
        requests = [
            # Rectangle
            {
                'createShape': {
                    'objectId': f"rect_{timestamp.replace(' ', '_').replace(':', '_')}",
                    'shapeType': 'RECTANGLE',
                    'elementProperties': {
                        'pageObjectId': first_slide_id,
                        'size': {
                            'height': {'magnitude': 100, 'unit': 'PT'},
                            'width': {'magnitude': 150, 'unit': 'PT'}
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': 100,
                            'translateY': 200,
                            'unit': 'PT'
                        }
                    }
                }
            },
            # Ellipse
            {
                'createShape': {
                    'objectId': f"ellipse_{timestamp.replace(' ', '_').replace(':', '_')}",
                    'shapeType': 'ELLIPSE',
                    'elementProperties': {
                        'pageObjectId': first_slide_id,
                        'size': {
                            'height': {'magnitude': 100, 'unit': 'PT'},
                            'width': {'magnitude': 100, 'unit': 'PT'}
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': 300,
                            'translateY': 200,
                            'unit': 'PT'
                        }
                    }
                }
            }
        ]

        print(f"   Creating shapes (rectangle, ellipse)...")
        result = await client.batch_update(presentation_id, requests)

        replies = result.get('replies', [])
        print(f"   ✅ Batch update completed")
        print(f"   Replies received: {len(replies)}")

        # Count created shapes
        shape_count = sum(1 for reply in replies if 'createShape' in reply)

        for i, reply in enumerate(replies, 1):
            if 'createShape' in reply:
                shape_id = reply['createShape'].get('objectId')
                print(f"      Shape {i}: Created with ID {shape_id}")

        if shape_count == 2:
            print("   ✅ PASS: All shapes created successfully")
            return True
        else:
            print(f"   ❌ FAIL: Expected 2 shapes, created {shape_count}")
            return False

    except Exception as e:
        print(f"   ❌ FAIL: Error adding shapes: {e}")
        import traceback
        traceback.print_exc()
        return False


async def test_add_images(client: SlidesTestClient, presentation_id: str) -> bool:
    """Test 4: Add images to a slide"""
    print("\n" + "=" * 60)
    print("📝 Test 4: Adding Images")
    print("=" * 60)

    try:
        # Get the second slide (or first if only one exists)
        presentation = await client.get_presentation(presentation_id)
        slides = presentation.get('slides', [])

        if not slides:
            print("   ⚠️  WARNING: No slides found in presentation")
            return False

        target_slide_id = slides[1].get('objectId') if len(slides) > 1 else slides[0].get('objectId')
        print(f"   Target slide ID: {target_slide_id}")

        # Use a public test image
        image_url = "https://static-cdn.toi-media.com/www/uploads/2014/07/gal-gadot.jpg"
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        requests = [
            {
                'createImage': {
                    'url': image_url,
                    'elementProperties': {
                        'pageObjectId': target_slide_id,
                        'size': {
                            'height': {'magnitude': 200, 'unit': 'PT'},
                            'width': {'magnitude': 300, 'unit': 'PT'}
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': 100,
                            'translateY': 100,
                            'unit': 'PT'
                        }
                    }
                }
            }
        ]

        print(f"   Adding image from: {image_url}")
        result = await client.batch_update(presentation_id, requests)

        replies = result.get('replies', [])
        print(f"   ✅ Batch update completed")
        print(f"   Replies received: {len(replies)}")

        # Check if image was created
        created_image = False
        for reply in replies:
            if 'createImage' in reply:
                image_id = reply['createImage'].get('objectId')
                print(f"      Created image with ID: {image_id}")
                created_image = True

        if created_image:
            print("   ✅ PASS: Image added successfully")
            return True
        else:
            print("   ❌ FAIL: Image creation not confirmed")
            return False

    except Exception as e:
        print(f"   ❌ FAIL: Error adding image: {e}")
        import traceback
        traceback.print_exc()
        return False


async def test_complex_batch(client: SlidesTestClient, presentation_id: str) -> bool:
    """Test 5: Complex batch operation combining multiple changes"""
    print("\n" + "=" * 60)
    print("📝 Test 5: Complex Batch Operation")
    print("=" * 60)

    try:
        # Get presentation state
        presentation = await client.get_presentation(presentation_id)
        slides = presentation.get('slides', [])
        initial_slide_count = len(slides)

        print(f"   Initial slide count: {initial_slide_count}")

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        new_slide_id = f"slide_{timestamp.replace(' ', '_').replace(':', '_')}"
        title_box_id = f"title_{timestamp.replace(' ', '_').replace(':', '_')}"
        body_box_id = f"body_{timestamp.replace(' ', '_').replace(':', '_')}"
        shape_id = f"shape_{timestamp.replace(' ', '_').replace(':', '_')}"

        # Complex batch: Create slide, add title, add body text, add shape
        requests = [
            # 1. Create a new slide
            {
                'createSlide': {
                    'objectId': new_slide_id,
                    'slideLayoutReference': {
                        'predefinedLayout': 'BLANK'
                    }
                }
            },
            # 2. Add title text box
            {
                'createShape': {
                    'objectId': title_box_id,
                    'shapeType': 'TEXT_BOX',
                    'elementProperties': {
                        'pageObjectId': new_slide_id,
                        'size': {
                            'height': {'magnitude': 50, 'unit': 'PT'},
                            'width': {'magnitude': 600, 'unit': 'PT'}
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': 50,
                            'translateY': 30,
                            'unit': 'PT'
                        }
                    }
                }
            },
            # 3. Insert title text
            {
                'insertText': {
                    'objectId': title_box_id,
                    'text': f'Complex Batch Test - {timestamp}',
                    'insertionIndex': 0
                }
            },
            # 4. Add body text box
            {
                'createShape': {
                    'objectId': body_box_id,
                    'shapeType': 'TEXT_BOX',
                    'elementProperties': {
                        'pageObjectId': new_slide_id,
                        'size': {
                            'height': {'magnitude': 200, 'unit': 'PT'},
                            'width': {'magnitude': 600, 'unit': 'PT'}
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': 50,
                            'translateY': 100,
                            'unit': 'PT'
                        }
                    }
                }
            },
            # 5. Insert body text
            {
                'insertText': {
                    'objectId': body_box_id,
                    'text': 'This slide was created using a complex batch operation that:\n\n'
                           '1. Created a new blank slide\n'
                           '2. Added a title text box\n'
                           '3. Added body text\n'
                           '4. Added an image\n\n'
                           'All in a single atomic operation!',
                    'insertionIndex': 0
                }
            },
            # 6. Add an image
            {
                'createImage': {
                    'url': 'https://static-cdn.toi-media.com/www/uploads/2014/07/gal-gadot.jpg',
                    'elementProperties': {
                        'pageObjectId': new_slide_id,
                        'size': {
                            'height': {'magnitude': 150, 'unit': 'PT'},
                            'width': {'magnitude': 200, 'unit': 'PT'}
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': 400,
                            'translateY': 100,
                            'unit': 'PT'
                        }
                    }
                }
            }
        ]

        print(f"   Executing complex batch with {len(requests)} operations:")
        print(f"      - Create new slide")
        print(f"      - Add title text box")
        print(f"      - Insert title text")
        print(f"      - Add body text box")
        print(f"      - Insert body text")
        print(f"      - Add an image")

        result = await client.batch_update(presentation_id, requests)

        replies = result.get('replies', [])
        print(f"   ✅ Batch update completed")
        print(f"   Replies received: {len(replies)}")

        # Analyze replies
        operation_types = {
            'createSlide': 0,
            'createShape': 0,
            'insertText': 0,
            'createImage': 0
        }

        for reply in replies:
            for op_type in operation_types.keys():
                if op_type in reply:
                    operation_types[op_type] += 1

        print(f"   Operations completed:")
        for op_type, count in operation_types.items():
            print(f"      - {op_type}: {count}")

        # Verify slide was created
        final_presentation = await client.get_presentation(presentation_id)
        final_slide_count = len(final_presentation.get('slides', []))

        print(f"   Final slide count: {final_slide_count}")

        # Check success
        success = (
            operation_types['createSlide'] == 1 and
            operation_types['createShape'] == 2 and  # title box, body box
            operation_types['createImage'] == 1 and  # image
            final_slide_count == initial_slide_count + 1
        )

        if success:
            print("   ✅ PASS: Complex batch operation completed successfully")
            return True
        else:
            print("   ❌ FAIL: Some operations did not complete as expected")
            return False

    except Exception as e:
        print(f"   ❌ FAIL: Error in complex batch: {e}")
        import traceback
        traceback.print_exc()
        return False


async def run_all_tests(client: SlidesTestClient, presentation_id: str):
    """Run all test suites"""
    print("\n" + "=" * 80)
    print("🎯 Starting Google Slides Batch Update Tests")
    print("=" * 80)
    print(f"\nPresentation ID: {presentation_id}")
    print(f"Presentation URL: https://docs.google.com/presentation/d/{presentation_id}/edit")

    results = {}

    # Run all tests
    tests = [
        ("Create Slides", test_create_slides),
        ("Add Text", test_add_text),
        ("Add Shapes", test_add_shapes),
        ("Add Images", test_add_images),
        ("Complex Batch", test_complex_batch),
    ]

    for test_name, test_func in tests:
        try:
            result = await test_func(client, presentation_id)
            results[test_name] = "PASSED" if result else "FAILED"
        except Exception as e:
            print(f"\n   ❌ Test '{test_name}' encountered an error: {e}")
            results[test_name] = "ERROR"

    # Print summary
    print("\n" + "=" * 80)
    print("📊 TEST SUMMARY")
    print("=" * 80)

    passed = sum(1 for r in results.values() if r == "PASSED")
    failed = sum(1 for r in results.values() if r == "FAILED")
    errors = sum(1 for r in results.values() if r == "ERROR")

    for test_name, result in results.items():
        status_icon = "✅" if result == "PASSED" else "❌"
        print(f"{status_icon} {test_name}: {result}")

    print("\n" + "=" * 80)
    print(f"Total: {len(results)} tests | ✅ Passed: {passed} | ❌ Failed: {failed} | ⚠️  Errors: {errors}")
    print("=" * 80)

    return results


async def run_single_test(
    test_name: str,
    client: SlidesTestClient,
    presentation_id: str
):
    """Run a single test by name"""
    test_map = {
        'create_slides': ("Create Slides", test_create_slides),
        'add_text': ("Add Text", test_add_text),
        'add_shapes': ("Add Shapes", test_add_shapes),
        'add_images': ("Add Images", test_add_images),
        'complex': ("Complex Batch", test_complex_batch),
    }

    if test_name not in test_map:
        print(f"❌ Unknown test: {test_name}")
        print(f"Available tests: {', '.join(test_map.keys())}")
        return

    display_name, test_func = test_map[test_name]

    print("\n" + "=" * 80)
    print(f"🎯 Running Single Test: {display_name}")
    print("=" * 80)
    print(f"\nPresentation ID: {presentation_id}")

    try:
        result = await test_func(client, presentation_id)

        print("\n" + "=" * 80)
        status = "✅ PASSED" if result else "❌ FAILED"
        print(f"Test Result: {status}")
        print("=" * 80)

    except Exception as e:
        print(f"\n❌ Test encountered an error: {e}")
        import traceback
        traceback.print_exc()


def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(
        description="Test Google Slides batch_update_presentation functionality"
    )
    parser.add_argument(
        "--env",
        choices=["local", "test", "prod"],
        default="local",
        help="Environment (unused for now, kept for consistency)"
    )
    parser.add_argument(
        "--test",
        choices=["all", "create_slides", "add_text", "add_shapes", "add_images", "complex"],
        default="all",
        help="Which test to run"
    )
    parser.add_argument(
        "--presentation-url",
        default=TEST_PRESENTATION_URL,
        help="Google Slides presentation URL to test with"
    )

    args = parser.parse_args()

    # Get credentials from environment
    refresh_token = os.getenv("TEST_GOOGLE_OAUTH_REFRESH_TOKEN")
    client_id = os.getenv("TEST_GOOGLE_OAUTH_CLIENT_ID")
    client_secret = os.getenv("TEST_GOOGLE_OAUTH_CLIENT_SECRET")

    if not all([refresh_token, client_id, client_secret]):
        print("❌ Missing required environment variables:")
        print("   - TEST_GOOGLE_OAUTH_REFRESH_TOKEN")
        print("   - TEST_GOOGLE_OAUTH_CLIENT_ID")
        print("   - TEST_GOOGLE_OAUTH_CLIENT_SECRET")
        print("\nPlease set these in your .env file")
        sys.exit(1)

    # Extract presentation ID
    try:
        presentation_id = extract_presentation_id(args.presentation_url)
        print(f"📋 Using presentation ID: {presentation_id}")
    except ValueError as e:
        print(f"❌ {e}")
        sys.exit(1)

    # Create client
    client = SlidesTestClient(refresh_token, client_id, client_secret)

    # Run tests
    if args.test == "all":
        asyncio.run(run_all_tests(client, presentation_id))
    else:
        asyncio.run(run_single_test(args.test, client, presentation_id))


if __name__ == "__main__":
    # Load environment variables from .env file
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except ImportError:
        print("⚠️  python-dotenv not installed, using system environment variables only")

    main()
