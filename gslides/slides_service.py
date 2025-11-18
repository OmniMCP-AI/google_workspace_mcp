"""
Google Slides Service Helper Functions

This module provides helper functions for Google Slides operations.
"""

import logging
import asyncio
from typing import Dict, Any, Optional, Tuple
from googleapiclient.errors import HttpError

logger = logging.getLogger(__name__)

# Counter for sequential object IDs
_object_id_counters = {
    'slide': 0,
    'title': 0,
    'body': 0,
    'image': 0
}


def generate_object_id(prefix: str) -> str:
    """
    Generate a sequential object ID with the given prefix.

    Args:
        prefix: The prefix for the object ID (e.g., 'slide', 'title', 'body', 'image')

    Returns:
        A unique object ID like 'title_1', 'body_2', etc.
    """
    _object_id_counters[prefix] += 1
    return f"{prefix}_{_object_id_counters[prefix]}"


async def get_slide_by_position(
    service,
    presentation_id: str,
    position: str
) -> Optional[Tuple[str, int]]:
    """
    Get slide ID by position ('first', 'last', or specific slide_id).

    Args:
        service: The Google Slides service instance
        presentation_id: The presentation ID
        position: 'first', 'last', or a specific slide object ID

    Returns:
        Tuple of (slide_id, total_slides) or None if not found

    Raises:
        ValueError: If position is invalid or slide not found
    """
    # Get the presentation
    result = await asyncio.to_thread(
        service.presentations().get(presentationId=presentation_id).execute
    )

    slides = result.get('slides', [])

    if not slides:
        raise ValueError(f"Presentation has no slides. Create a slide first.")

    # Handle position
    if position == "first":
        return slides[0].get('objectId'), len(slides)
    elif position == "last":
        return slides[-1].get('objectId'), len(slides)
    else:
        # Check if the specific slide_id exists
        slide_ids = [slide.get('objectId') for slide in slides]
        if position in slide_ids:
            return position, len(slides)
        else:
            available_ids = ', '.join(slide_ids[:3])
            if len(slide_ids) > 3:
                available_ids += f", ... ({len(slide_ids)} total)"
            raise ValueError(
                f"Slide '{position}' not found. "
                f"Presentation has {len(slides)} slide(s). "
                f"Available slide IDs: {available_ids}"
            )


async def extract_presentation_id_from_url(presentation_url: str) -> str:
    """
    Extract presentation ID from a Google Slides URL.

    Args:
        presentation_url: URL or ID of the presentation

    Returns:
        The presentation ID
    """
    # If it's already just an ID (no slashes), return it
    if '/' not in presentation_url:
        return presentation_url

    # Extract from URL: https://docs.google.com/presentation/d/{id}/edit
    if '/presentation/d/' in presentation_url:
        parts = presentation_url.split('/presentation/d/')
        if len(parts) > 1:
            presentation_id = parts[1].split('/')[0]
            # Remove any URL fragments or query params
            presentation_id = presentation_id.split('?')[0].split('#')[0]
            return presentation_id

    # If we can't parse it, assume it's an ID
    return presentation_url


def create_text_box_request(
    object_id: str,
    page_id: str,
    text: str,
    size_height: int,
    size_width: int,
    x: int,
    y: int
) -> list:
    """
    Create requests for adding a text box to a slide.

    Returns a list of 2 requests: createShape and insertText
    """
    return [
        {
            'createShape': {
                'objectId': object_id,
                'shapeType': 'TEXT_BOX',
                'elementProperties': {
                    'pageObjectId': page_id,
                    'size': {
                        'height': {'magnitude': size_height, 'unit': 'PT'},
                        'width': {'magnitude': size_width, 'unit': 'PT'}
                    },
                    'transform': {
                        'scaleX': 1,
                        'scaleY': 1,
                        'translateX': x,
                        'translateY': y,
                        'unit': 'PT'
                    }
                }
            }
        },
        {
            'insertText': {
                'objectId': object_id,
                'text': text,
                'insertionIndex': 0
            }
        }
    ]


def create_image_request(
    object_id: str,
    page_id: str,
    image_url: str,
    size_height: int,
    size_width: int,
    x: int,
    y: int
) -> dict:
    """
    Create request for adding an image to a slide.
    """
    return {
        'createImage': {
            'objectId': object_id,
            'url': image_url,
            'elementProperties': {
                'pageObjectId': page_id,
                'size': {
                    'height': {'magnitude': size_height, 'unit': 'PT'},
                    'width': {'magnitude': size_width, 'unit': 'PT'}
                },
                'transform': {
                    'scaleX': 1,
                    'scaleY': 1,
                    'translateX': x,
                    'translateY': y,
                    'unit': 'PT'
                }
            }
        }
    }
