"""
Google Slides MCP Tools

This module provides MCP tools for interacting with Google Slides API.
"""

import logging
import asyncio
from typing import List, Optional, Dict, Any

from mcp import types
from fastmcp import Context
from googleapiclient.errors import HttpError

from auth.service_decorator import require_google_service
from core.server import server
from core.utils import handle_http_errors
from core.comments import create_comment_tools
from gslides.slides_service import (
    generate_object_id,
    get_slide_by_position,
    extract_presentation_id_from_url,
    create_text_box_request,
    create_image_request
)
from gslides.slides_models import (
    CreatePresentationResponse,
    AddSlideResponse,
    AddTitleResponse,
    AddBodyTextResponse,
    AddBodyImageResponse,
    AddPageWithContentResponse,
    MarkdownToSlidesResponse,
    ErrorResponse
)
from gslides.markdown_parser import parse_markdown_to_slides

logger = logging.getLogger(__name__)


@server.tool
@require_google_service("slides", "slides")
@handle_http_errors("create_presentation")
async def create_presentation(
    service,
    ctx: Context,
    user_google_email: Optional[str] = None,
    title: str = "Untitled Presentation"
) -> CreatePresentationResponse:
    """
    <description>Creates a new empty Google Slides presentation with a single blank slide. Generates a presentation ready for content addition but contains no slides content initially.</description>
    
    <use_case>Starting new presentations for meetings, reports, or educational content. Ideal when building presentations from scratch rather than using templates.</use_case>
    
    <limitation>Creates only one blank slide - use batch_update_presentation to add more slides or content. Cannot import slides from other presentations or apply themes during creation.</limitation>
    
    <failure_cases>Fails if user lacks Google Slides creation permissions, if title exceeds character limits, or if Google Drive storage quota is exceeded.</failure_cases>

    Args:
        user_google_email (Optional[str]): The user's Google email address. If not provided, will be automatically detected.
        title (str): The title for the new presentation. Defaults to "Untitled Presentation".

    Returns:
        str: Details about the created presentation including ID and URL.
    """
    logger.info(f"[create_presentation] Invoked. Email: '{user_google_email}', Title: '{title}'")

    body = {
        'title': title
    }
    
    result = await asyncio.to_thread(
        service.presentations().create(body=body).execute
    )

    presentation_id = result.get('presentationId')
    presentation_url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"
    slides = result.get('slides', [])

    logger.info(f"Presentation created successfully for {user_google_email}")
    return CreatePresentationResponse(
        success=True,
        presentation_id=presentation_id,
        presentation_url=presentation_url,
        title=title,
        slides_created=len(slides)
    )


@server.tool
@require_google_service("slides", "slides_read")
@handle_http_errors("get_presentation")
async def get_presentation(
    service,
    ctx: Context,
    presentation_id: str,
    user_google_email: Optional[str] = None
):
    """
    <description>Retrieves presentation metadata including title, slide count, page dimensions, and basic slide structure. Shows presentation overview without detailed slide content.</description>
    
    <use_case>Inspecting presentation structure before modification, understanding slide organization for automation, or getting presentation metadata for documentation.</use_case>
    
    <limitation>Returns slide structure but not detailed content like text or images. Cannot retrieve presentations without proper access permissions or deleted presentations.</limitation>
    
    <failure_cases>Fails with invalid presentation IDs, presentations the user cannot access due to sharing restrictions, or presentations deleted by the owner.</failure_cases>

    Args:
        user_google_email (Optional[str]): The user's Google email address. If not provided, will be automatically detected.
        presentation_id (str): The ID of the presentation to retrieve.

    Returns:
        str: Details about the presentation including title, slides count, and metadata.
    """
    logger.info(f"[get_presentation] Invoked. Email: '{user_google_email}', ID: '{presentation_id}'")

    result = await asyncio.to_thread(
        service.presentations().get(presentationId=presentation_id).execute
    )
    
    title = result.get('title', 'Untitled')
    slides = result.get('slides', [])
    page_size = result.get('pageSize', {})
    
    slides_info = []
    for i, slide in enumerate(slides, 1):
        slide_id = slide.get('objectId', 'Unknown')
        page_elements = slide.get('pageElements', [])
        slides_info.append(f"  Slide {i}: ID {slide_id}, {len(page_elements)} element(s)")
    
    confirmation_message = f"""Presentation Details for {user_google_email}:
- Title: {title}
- Presentation ID: {presentation_id}
- URL: https://docs.google.com/presentation/d/{presentation_id}/edit
- Total Slides: {len(slides)}
- Page Size: {page_size.get('width', {}).get('magnitude', 'Unknown')} x {page_size.get('height', {}).get('magnitude', 'Unknown')} {page_size.get('width', {}).get('unit', '')}

Slides Breakdown:
{chr(10).join(slides_info) if slides_info else '  No slides found'}"""
    
    logger.info(f"Presentation retrieved successfully for {user_google_email}")
    return confirmation_message


@server.tool
@require_google_service("slides", "slides")
@handle_http_errors("batch_update_presentation")
async def batch_update_presentation(
    service,
    ctx: Context,
    presentation_id: str,
    requests: List[Dict[str, Any]],
    user_google_email: Optional[str] = None
):
    """
    <description>Executes multiple presentation modifications in a single atomic operation including adding slides, inserting text, creating shapes, and applying formatting. All changes succeed or fail together.</description>
    
    <use_case>Automating slide creation, bulk formatting changes, inserting multiple elements across slides, or applying consistent styling to entire presentations programmatically.</use_case>
    
    <limitation>Requires knowledge of Google Slides API request structure. Limited to 500 requests per batch. Cannot undo individual operations - entire batch must be reverted.</limitation>
    
    <failure_cases>Fails if any single request in the batch is invalid, if user lacks edit permissions, or if presentation is locked by another user during the operation.</failure_cases>

    Args:
        user_google_email (Optional[str]): The user's Google email address. If not provided, will be automatically detected.
        presentation_id (str): The ID of the presentation to update.
        requests (List[Dict[str, Any]]): List of update requests to apply.

    Returns:
        str: Details about the batch update operation results.
    """
    logger.info(f"[batch_update_presentation] Invoked. Email: '{user_google_email}', ID: '{presentation_id}', Requests: {len(requests)}")

    body = {
        'requests': requests
    }
    
    result = await asyncio.to_thread(
        service.presentations().batchUpdate(
            presentationId=presentation_id,
            body=body
        ).execute
    )
    
    replies = result.get('replies', [])
    
    confirmation_message = f"""Batch Update Completed for {user_google_email}:
- Presentation ID: {presentation_id}
- URL: https://docs.google.com/presentation/d/{presentation_id}/edit
- Requests Applied: {len(requests)}
- Replies Received: {len(replies)}"""
    
    if replies:
        confirmation_message += "\n\nUpdate Results:"
        for i, reply in enumerate(replies, 1):
            if 'createSlide' in reply:
                slide_id = reply['createSlide'].get('objectId', 'Unknown')
                confirmation_message += f"\n  Request {i}: Created slide with ID {slide_id}"
            elif 'createShape' in reply:
                shape_id = reply['createShape'].get('objectId', 'Unknown')
                confirmation_message += f"\n  Request {i}: Created shape with ID {shape_id}"
            else:
                confirmation_message += f"\n  Request {i}: Operation completed"
    
    logger.info(f"Batch update completed successfully for {user_google_email}")
    return confirmation_message


@server.tool
@require_google_service("slides", "slides_read")
@handle_http_errors("get_page")
async def get_page(
    service,
    ctx: Context,
    presentation_id: str,
    page_object_id: str,
    user_google_email: Optional[str] = None
):
    """
    <description>Retrieves detailed information about a specific slide including all page elements (text boxes, shapes, images, tables) and their properties. Shows slide content structure.</description>
    
    <use_case>Analyzing slide content before modification, understanding element layout for automation, or extracting specific slide information for processing.</use_case>
    
    <limitation>Returns element metadata but not actual text content or image data. Cannot retrieve slides from presentations without proper access permissions.</limitation>
    
    <failure_cases>Fails with invalid presentation or slide IDs, slides the user cannot access, or slides that have been deleted from the presentation.</failure_cases>

    Args:
        user_google_email (Optional[str]): The user's Google email address. If not provided, will be automatically detected.
        presentation_id (str): The ID of the presentation.
        page_object_id (str): The object ID of the page/slide to retrieve.

    Returns:
        str: Details about the specific page including elements and layout.
    """
    logger.info(f"[get_page] Invoked. Email: '{user_google_email}', Presentation: '{presentation_id}', Page: '{page_object_id}'")

    result = await asyncio.to_thread(
        service.presentations().pages().get(
            presentationId=presentation_id,
            pageObjectId=page_object_id
        ).execute
    )
    
    page_type = result.get('pageType', 'Unknown')
    page_elements = result.get('pageElements', [])
    
    elements_info = []
    for element in page_elements:
        element_id = element.get('objectId', 'Unknown')
        if 'shape' in element:
            shape_type = element['shape'].get('shapeType', 'Unknown')
            elements_info.append(f"  Shape: ID {element_id}, Type: {shape_type}")
        elif 'table' in element:
            table = element['table']
            rows = table.get('rows', 0)
            cols = table.get('columns', 0)
            elements_info.append(f"  Table: ID {element_id}, Size: {rows}x{cols}")
        elif 'line' in element:
            line_type = element['line'].get('lineType', 'Unknown')
            elements_info.append(f"  Line: ID {element_id}, Type: {line_type}")
        else:
            elements_info.append(f"  Element: ID {element_id}, Type: Unknown")
    
    confirmation_message = f"""Page Details for {user_google_email}:
- Presentation ID: {presentation_id}
- Page ID: {page_object_id}
- Page Type: {page_type}
- Total Elements: {len(page_elements)}

Page Elements:
{chr(10).join(elements_info) if elements_info else '  No elements found'}"""
    
    logger.info(f"Page retrieved successfully for {user_google_email}")
    return confirmation_message


@server.tool
@require_google_service("slides", "slides_read")
@handle_http_errors("get_page_thumbnail")
async def get_page_thumbnail(
    service,
    ctx: Context,
    presentation_id: str,
    page_object_id: str,
    user_google_email: Optional[str] = None,
    thumbnail_size: str = "MEDIUM"
):
    """
    <description>Generates a downloadable thumbnail image URL for a specific slide in PNG format. Creates visual preview of slide content at specified resolution (SMALL: 200px, MEDIUM: 800px, LARGE: 1600px width).</description>
    
    <use_case>Creating slide previews for presentation catalogs, generating thumbnails for slide selection interfaces, or extracting slide images for documentation or reports.</use_case>
    
    <limitation>Returns temporary URLs that expire after some time. Cannot generate thumbnails for slides with restricted content or presentations without view permissions.</limitation>
    
    <failure_cases>Fails with invalid presentation or slide IDs, slides containing restricted content the user cannot view, or when Google's thumbnail generation service is temporarily unavailable.</failure_cases>

    Args:
        user_google_email (Optional[str]): The user's Google email address. If not provided, will be automatically detected.
        presentation_id (str): The ID of the presentation.
        page_object_id (str): The object ID of the page/slide.
        thumbnail_size (str): Size of thumbnail ("LARGE", "MEDIUM", "SMALL"). Defaults to "MEDIUM".

    Returns:
        str: URL to the generated thumbnail image.
    """
    logger.info(f"[get_page_thumbnail] Invoked. Email: '{user_google_email}', Presentation: '{presentation_id}', Page: '{page_object_id}', Size: '{thumbnail_size}'")

    result = await asyncio.to_thread(
        service.presentations().pages().getThumbnail(
            presentationId=presentation_id,
            pageObjectId=page_object_id,
            thumbnailPropertiesImageSize=thumbnail_size
        ).execute
    )
    
    thumbnail_url = result.get('contentUrl', '')
    
    confirmation_message = f"""Thumbnail Generated for {user_google_email}:
- Presentation ID: {presentation_id}
- Page ID: {page_object_id}
- Thumbnail Size: {thumbnail_size}
- Thumbnail URL: {thumbnail_url}

You can view or download the thumbnail using the provided URL."""
    
    logger.info(f"Thumbnail generated successfully for {user_google_email}")
    return confirmation_message


@server.tool
@require_google_service("slides", "slides")
@handle_http_errors("add_slide")
async def add_slide(
    service,
    ctx: Context,
    presentation_url: str,
    layout: str = "BLANK",
    slide_id: Optional[str] = None,
    user_google_email: Optional[str] = None
) -> AddSlideResponse:
    """
    <description>Adds a new slide to an existing presentation with the specified layout (BLANK, TITLE_AND_BODY, etc.). Simplifies slide creation with optional custom slide ID.</description>

    <use_case>Adding slides to presentations programmatically, building multi-slide decks, creating presentation templates with predefined layouts.</use_case>

    <limitation>Layout must be a valid Google Slides predefined layout name. Cannot add slides to presentations without edit permissions.</limitation>

    <failure_cases>Fails with invalid presentation URLs, unsupported layout names, or when user lacks edit permissions on the presentation.</failure_cases>

    Args:
        presentation_url (str): The URL or ID of the presentation
        layout (str): Slide layout (BLANK, TITLE_AND_BODY, TITLE_ONLY, etc.). Defaults to "BLANK"
        slide_id (Optional[str]): Custom slide ID. Auto-generated if not provided
        user_google_email (Optional[str]): The user's Google email address

    Returns:
        str: JSON with slide_id and presentation_id
    """
    logger.info(f"[add_slide] Invoked. URL: '{presentation_url}', Layout: '{layout}'")

    try:
        presentation_id = await extract_presentation_id_from_url(presentation_url)

        # Generate slide ID if not provided
        if not slide_id:
            slide_id = generate_object_id('slide')

        # Create the slide request
        requests = [
            {
                'createSlide': {
                    'objectId': slide_id,
                    'slideLayoutReference': {
                        'predefinedLayout': layout
                    }
                }
            }
        ]

        # Execute the batch update
        await asyncio.to_thread(
            service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': requests}
            ).execute
        )

        logger.info(f"Slide added successfully: {slide_id}")
        return AddSlideResponse(
            success=True,
            slide_id=slide_id,
            presentation_id=presentation_id,
            layout=layout
        )

    except ValueError as e:
        logger.error(f"Validation error in add_slide: {e}")
        return ErrorResponse(success=False, error=str(e))


@server.tool
@require_google_service("slides", "slides")
@handle_http_errors("add_title")
async def add_title(
    service,
    ctx: Context,
    presentation_url: str,
    text: str,
    page_id: str = "last",
    size_height: int = 50,
    size_width: int = 600,
    x: int = 50,
    y: int = 50,
    user_google_email: Optional[str] = None
) -> AddTitleResponse:
    """
    <description>Adds a title text box to a specific slide with customizable position and size. Defaults to adding to the last slide in the presentation.</description>

    <use_case>Adding titles to slides programmatically, creating consistent slide headers, automating presentation generation with standardized formatting.</use_case>

    <limitation>Requires valid slide ID or "first"/"last" position. Cannot add titles to slides without edit permissions. Text formatting is basic.</limitation>

    <failure_cases>Fails when page_id doesn't exist, when presentation has no slides, or when user lacks edit permissions.</failure_cases>

    Args:
        presentation_url (str): The URL or ID of the presentation
        text (str): The title text to add
        page_id (str): Slide ID, "first", or "last". Defaults to "last"
        size_height (int): Height in points. Defaults to 50
        size_width (int): Width in points. Defaults to 600
        x (int): X position in points. Defaults to 50
        y (int): Y position in points. Defaults to 50
        user_google_email (Optional[str]): The user's Google email address

    Returns:
        str: JSON with title_object_id and slide_id
    """
    logger.info(f"[add_title] Invoked. URL: '{presentation_url}', Text: '{text[:50]}...', Page: '{page_id}'")

    try:
        presentation_id = await extract_presentation_id_from_url(presentation_url)

        # Resolve the slide ID
        slide_id, total_slides = await get_slide_by_position(service, presentation_id, page_id)

        # Generate title object ID
        title_id = generate_object_id('title')

        # Create requests for text box
        requests = create_text_box_request(
            object_id=title_id,
            page_id=slide_id,
            text=text,
            size_height=size_height,
            size_width=size_width,
            x=x,
            y=y
        )

        # Execute the batch update
        await asyncio.to_thread(
            service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': requests}
            ).execute
        )

        logger.info(f"Title added successfully: {title_id}")
        return AddTitleResponse(
            success=True,
            title_object_id=title_id,
            slide_id=slide_id,
            presentation_id=presentation_id
        )

    except ValueError as e:
        logger.error(f"Validation error in add_title: {e}")
        return ErrorResponse(success=False, error=str(e))


@server.tool
@require_google_service("slides", "slides")
@handle_http_errors("add_body_text")
async def add_body_text(
    service,
    ctx: Context,
    presentation_url: str,
    text: str,
    page_id: str = "last",
    size_height: int = 200,
    size_width: int = 600,
    x: int = 50,
    y: int = 100,
    user_google_email: Optional[str] = None
) -> AddBodyTextResponse:
    """
    <description>Adds body text to a specific slide with customizable position and size. Defaults to adding below typical title position on the last slide.</description>

    <use_case>Adding content to slides programmatically, creating slide descriptions, automating presentation content generation with consistent formatting.</use_case>

    <limitation>Requires valid slide ID or "first"/"last" position. Cannot add text to slides without edit permissions. Text formatting is basic.</limitation>

    <failure_cases>Fails when page_id doesn't exist, when presentation has no slides, or when user lacks edit permissions.</failure_cases>

    Args:
        presentation_url (str): The URL or ID of the presentation
        text (str): The body text to add
        page_id (str): Slide ID, "first", or "last". Defaults to "last"
        size_height (int): Height in points. Defaults to 200
        size_width (int): Width in points. Defaults to 600
        x (int): X position in points. Defaults to 50
        y (int): Y position in points. Defaults to 100
        user_google_email (Optional[str]): The user's Google email address

    Returns:
        str: JSON with body_object_id and slide_id
    """
    logger.info(f"[add_body_text] Invoked. URL: '{presentation_url}', Text: '{text[:50]}...', Page: '{page_id}'")

    try:
        presentation_id = await extract_presentation_id_from_url(presentation_url)

        # Resolve the slide ID
        slide_id, total_slides = await get_slide_by_position(service, presentation_id, page_id)

        # Generate body object ID
        body_id = generate_object_id('body')

        # Create requests for text box
        requests = create_text_box_request(
            object_id=body_id,
            page_id=slide_id,
            text=text,
            size_height=size_height,
            size_width=size_width,
            x=x,
            y=y
        )

        # Execute the batch update
        await asyncio.to_thread(
            service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': requests}
            ).execute
        )

        logger.info(f"Body text added successfully: {body_id}")
        return AddBodyTextResponse(
            success=True,
            body_object_id=body_id,
            slide_id=slide_id,
            presentation_id=presentation_id
        )

    except ValueError as e:
        logger.error(f"Validation error in add_body_text: {e}")
        return ErrorResponse(success=False, error=str(e))


@server.tool
@require_google_service("slides", "slides")
@handle_http_errors("add_body_image")
async def add_body_image(
    service,
    ctx: Context,
    presentation_url: str,
    image_url: str,
    page_id: str = "last",
    size_height: int = 150,
    size_width: int = 200,
    x: int = 100,
    y: int = 50,
    user_google_email: Optional[str] = None
) -> AddBodyImageResponse:
    """
    <description>Adds an image to a specific slide from a URL with customizable position and size. Defaults to adding to the last slide in the presentation.</description>

    <use_case>Adding images to slides programmatically, inserting logos or diagrams, automating visual content creation in presentations.</use_case>

    <limitation>Requires publicly accessible image URL. Cannot add images to slides without edit permissions. Image must be in supported format (PNG, JPG, GIF).</limitation>

    <failure_cases>Fails when page_id doesn't exist, when image URL is invalid or inaccessible, when presentation has no slides, or when user lacks edit permissions.</failure_cases>

    Args:
        presentation_url (str): The URL or ID of the presentation
        image_url (str): The URL of the image to add
        page_id (str): Slide ID, "first", or "last". Defaults to "last"
        size_height (int): Height in points. Defaults to 150
        size_width (int): Width in points. Defaults to 200
        x (int): X position in points. Defaults to 100
        y (int): Y position in points. Defaults to 50
        user_google_email (Optional[str]): The user's Google email address

    Returns:
        str: JSON with image_object_id and slide_id
    """
    logger.info(f"[add_body_image] Invoked. URL: '{presentation_url}', Image: '{image_url}', Page: '{page_id}'")

    try:
        presentation_id = await extract_presentation_id_from_url(presentation_url)

        # Resolve the slide ID
        slide_id, total_slides = await get_slide_by_position(service, presentation_id, page_id)

        # Generate image object ID
        image_id = generate_object_id('image')

        # Create request for image
        request = create_image_request(
            object_id=image_id,
            page_id=slide_id,
            image_url=image_url,
            size_height=size_height,
            size_width=size_width,
            x=x,
            y=y
        )

        # Execute the batch update
        await asyncio.to_thread(
            service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': [request]}
            ).execute
        )

        logger.info(f"Image added successfully: {image_id}")
        return AddBodyImageResponse(
            success=True,
            image_object_id=image_id,
            slide_id=slide_id,
            presentation_id=presentation_id
        )

    except ValueError as e:
        logger.error(f"Validation error in add_body_image: {e}")
        return ErrorResponse(success=False, error=str(e))


@server.tool
@require_google_service("slides", "slides")
@handle_http_errors("add_page_with_content")
async def add_page_with_content(
    service,
    ctx: Context,
    presentation_url: str,
    layout: str = "BLANK",
    title: Optional[str] = None,
    title_x: int = 30,
    title_y: int = 20,
    title_width: int = 640,
    title_height: int = 60,
    body_text: Optional[str] = None,
    body_x: int = 30,
    body_y: int = 100,
    body_width: int = 640,
    body_height: int = 300,
    image_url: Optional[str] = None,
    image_x: int = 400,
    image_y: int = 120,
    image_width: int = 250,
    image_height: int = 200,
    user_google_email: Optional[str] = None
) -> AddPageWithContentResponse:
    """
    <description>Creates a new slide with optional title, body text, and image in one atomic operation. This convenience wrapper combines add_slide, add_title, add_body_text, and add_body_image with customizable positioning.</description>

    <use_case>Rapidly creating complete slides with multiple elements, building presentation pages from templates, automating slide generation with consistent layouts. Perfect for batch slide creation or when you need title + content + image together.</use_case>

    <limitation>All elements are added with flat positioning parameters - no automatic layout management. If any element fails to add, previous elements remain (partial success). For fine-grained control over individual elements, use separate add_title, add_body_text, add_body_image tools.</limitation>

    <failure_cases>Fails with invalid presentation URLs, inaccessible image URLs, or when user lacks edit permissions. Partial failures possible if slide creates but elements fail to add.</failure_cases>

    Args:
        presentation_url (str): The URL or ID of the presentation
        layout (str): Slide layout (BLANK, TITLE_AND_BODY, etc.). Defaults to "BLANK"
        title (Optional[str]): Title text to add. If None, no title is added
        title_x (int): Title X position in points. Defaults to 30 (top-left)
        title_y (int): Title Y position in points. Defaults to 20 (top-left)
        title_width (int): Title width in points. Defaults to 640
        title_height (int): Title height in points. Defaults to 60
        body_text (Optional[str]): Body text to add. If None, no body is added
        body_x (int): Body X position in points. Defaults to 30 (aligned with title)
        body_y (int): Body Y position in points. Defaults to 100 (below title)
        body_width (int): Body width in points. Defaults to 640
        body_height (int): Body height in points. Defaults to 300
        image_url (Optional[str]): Image URL to add. If None, no image is added
        image_x (int): Image X position in points. Defaults to 400 (right side)
        image_y (int): Image Y position in points. Defaults to 120 (aligned with body)
        image_width (int): Image width in points. Defaults to 250
        image_height (int): Image height in points. Defaults to 200
        user_google_email (Optional[str]): The user's Google email address

    Returns:
        AddPageWithContentResponse: JSON with slide_id, presentation_id, layout, and list of elements_added
    """
    logger.info(f"[add_page_with_content] Invoked. URL: '{presentation_url}', Layout: '{layout}'")
    logger.info(f"  Title: {bool(title)}, Body: {bool(body_text)}, Image: {bool(image_url)}")

    try:
        presentation_id = await extract_presentation_id_from_url(presentation_url)

        # Generate IDs for all elements
        slide_id = generate_object_id('slide')
        title_id = generate_object_id('title') if title else None
        body_id = generate_object_id('body') if body_text else None
        image_id = generate_object_id('image') if image_url else None

        # Build requests array
        requests = []
        elements_added = []

        # Step 1: Create slide
        requests.append({
            'createSlide': {
                'objectId': slide_id,
                'slideLayoutReference': {
                    'predefinedLayout': layout
                }
            }
        })

        # Step 2: Add title if provided
        if title:
            title_requests = create_text_box_request(
                object_id=title_id,
                page_id=slide_id,
                text=title,
                size_height=title_height,
                size_width=title_width,
                x=title_x,
                y=title_y
            )
            requests.extend(title_requests)
            elements_added.append({
                "type": "title",
                "object_id": title_id
            })

        # Step 3: Add body text if provided
        if body_text:
            body_requests = create_text_box_request(
                object_id=body_id,
                page_id=slide_id,
                text=body_text,
                size_height=body_height,
                size_width=body_width,
                x=body_x,
                y=body_y
            )
            requests.extend(body_requests)
            elements_added.append({
                "type": "body_text",
                "object_id": body_id
            })

        # Step 4: Add image if provided
        if image_url:
            image_request = create_image_request(
                object_id=image_id,
                page_id=slide_id,
                image_url=image_url,
                size_height=image_height,
                size_width=image_width,
                x=image_x,
                y=image_y
            )
            requests.append(image_request)
            elements_added.append({
                "type": "image",
                "object_id": image_id
            })

        # Execute all requests in a single batch update
        await asyncio.to_thread(
            service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': requests}
            ).execute
        )

        logger.info(f"Page with content created successfully. Slide: {slide_id}, Elements: {len(elements_added)}")
        return AddPageWithContentResponse(
            success=True,
            slide_id=slide_id,
            presentation_id=presentation_id,
            layout=layout,
            elements_added=elements_added
        )

    except ValueError as e:
        logger.error(f"Validation error in add_page_with_content: {e}")
        return ErrorResponse(success=False, error=str(e))


@server.tool
@require_google_service("slides", "slides")
@handle_http_errors("create_presentation_from_markdown")
async def create_presentation_from_markdown(
    service,
    ctx: Context,
    markdown_content: str,
    presentation_title: Optional[str] = None,
    user_google_email: Optional[str] = None
):
    """
    <description>Creates a complete Google Slides presentation from Markdown content. Parses Markdown structure (H1=title, H2=slides, lists, images) and automatically generates slides with proper formatting.</description>

    <use_case>Rapidly converting documentation to presentations, creating slide decks from notes, automating report presentations, building educational content from Markdown. Perfect for content creators who write in Markdown.</use_case>

    <limitation>Requires H1 heading for presentation title (or provide presentation_title parameter). Only first image per slide is used. Tables are not supported. Code blocks converted to plain monospace text. H3+ headings become body text.</limitation>

    <failure_cases>Fails if markdown has no H1 and no presentation_title provided. Fails with invalid image URLs. Fails if user lacks Slides creation permissions. Returns warnings for unsupported elements (multiple images, tables).</failure_cases>

    Args:
        markdown_content (str): Markdown text content to convert
        presentation_title (Optional[str]): Override presentation title (instead of using H1)
        user_google_email (Optional[str]): The user's Google email address

    Markdown Parsing Rules:
        - # H1 → Presentation title (required, first H1 only)
        - ## H2 → New slide with title
        - ### H3 → Bold subheading in body text
        - Regular text → Body text
        - - Lists → Bullet points (converted to •)
        - 1. Lists → Numbered lists
        - ![alt](url) → Image on right side (first image only per slide)
        - ```code``` → Monospace code block
        - **bold** → Bold text formatting preserved
        - *italic* → Italic text formatting preserved

    Returns:
        MarkdownToSlidesResponse: JSON with presentation details, slides created, and warnings
    """
    logger.info(f"[create_presentation_from_markdown] Invoked. Content length: {len(markdown_content)}")

    try:
        # Parse markdown
        parsed_data = parse_markdown_to_slides(markdown_content)

        # Get title
        title = presentation_title or parsed_data.get("presentation_title")
        if not title:
            return ErrorResponse(
                success=False,
                error="No presentation title found. Markdown must start with # H1 heading, or provide presentation_title parameter."
            )

        slides_data = parsed_data.get("slides", [])
        warnings = parsed_data.get("warnings", [])

        if not slides_data:
            return ErrorResponse(
                success=False,
                error="No slides generated. Markdown must have at least one ## H2 heading to create slides."
            )

        logger.info(f"Parsed markdown: title='{title}', slides={len(slides_data)}, warnings={len(warnings)}")

        # Create presentation directly using API
        body = {'title': title}
        result = await asyncio.to_thread(
            service.presentations().create(body=body).execute
        )

        presentation_id = result.get('presentationId')
        presentation_url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"

        logger.info(f"Created presentation: {presentation_id}")

        # Create slides using add_page_with_content (directly build requests)
        total_elements = 0
        for slide_data in slides_data:
            try:
                # Extract presentation ID
                pres_id = await extract_presentation_id_from_url(presentation_url)

                # Generate IDs
                slide_id = generate_object_id('slide')
                title_id = generate_object_id('title') if slide_data.title else None
                body_id = generate_object_id('body') if slide_data.body_text else None
                image_id = generate_object_id('image') if slide_data.image_url else None

                # Build requests
                requests = []
                elements_count = 0

                # Create slide
                requests.append({
                    'createSlide': {
                        'objectId': slide_id,
                        'slideLayoutReference': {
                            'predefinedLayout': 'BLANK'
                        }
                    }
                })

                # Add title if exists
                if slide_data.title:
                    title_requests = create_text_box_request(
                        object_id=title_id,
                        page_id=slide_id,
                        text=slide_data.title,
                        size_height=60,
                        size_width=640,
                        x=30,
                        y=20
                    )
                    requests.extend(title_requests)
                    elements_count += 1

                # Add body text if exists
                if slide_data.body_text and slide_data.body_text.strip():
                    body_requests = create_text_box_request(
                        object_id=body_id,
                        page_id=slide_id,
                        text=slide_data.body_text.strip(),
                        size_height=300,
                        size_width=640,
                        x=30,
                        y=100
                    )
                    requests.extend(body_requests)
                    elements_count += 1

                # Add image if exists
                if slide_data.image_url:
                    image_request = create_image_request(
                        object_id=image_id,
                        page_id=slide_id,
                        image_url=slide_data.image_url,
                        size_height=200,
                        size_width=250,
                        x=400,
                        y=120
                    )
                    requests.append(image_request)
                    elements_count += 1

                # Execute batch update
                await asyncio.to_thread(
                    service.presentations().batchUpdate(
                        presentationId=pres_id,
                        body={'requests': requests}
                    ).execute
                )

                total_elements += elements_count
                logger.info(f"Created slide '{slide_data.title}' with {elements_count} elements")

            except Exception as e:
                logger.error(f"Error creating slide '{slide_data.title}': {e}")
                warnings.append(f"Error creating slide '{slide_data.title}': {str(e)}")

        logger.info(f"Presentation created successfully: {len(slides_data)} slides, {total_elements} elements")

        return MarkdownToSlidesResponse(
            success=True,
            presentation_id=presentation_id,
            presentation_url=presentation_url,
            presentation_title=title,
            slides_created=len(slides_data),
            total_elements=total_elements,
            warnings=warnings
        )

    except ValueError as e:
        logger.error(f"Validation error in create_presentation_from_markdown: {e}")
        return ErrorResponse(success=False, error=str(e))
    except Exception as e:
        logger.error(f"Unexpected error in create_presentation_from_markdown: {e}")
        return ErrorResponse(success=False, error=f"Unexpected error: {str(e)}")


# Create comment management tools for slides
_comment_tools = create_comment_tools("presentation", "presentation_id")
read_presentation_comments = _comment_tools['read_comments']
create_presentation_comment = _comment_tools['create_comment']
reply_to_presentation_comment = _comment_tools['reply_to_comment']
resolve_presentation_comment = _comment_tools['resolve_comment']

# Aliases for backwards compatibility and intuitive naming
read_slide_comments = read_presentation_comments
create_slide_comment = create_presentation_comment
reply_to_slide_comment = reply_to_presentation_comment
resolve_slide_comment = resolve_presentation_comment