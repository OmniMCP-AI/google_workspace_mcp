"""
Google Slides Response Models

Pydantic models for Google Slides API responses.
"""

from pydantic import BaseModel, Field


class CreatePresentationResponse(BaseModel):
    """Response from create_presentation"""
    success: bool = Field(description="Whether the operation succeeded")
    presentation_id: str = Field(description="The ID of the created presentation")
    presentation_url: str = Field(description="The URL to access the presentation")
    title: str = Field(description="The title of the presentation")
    slides_created: int = Field(description="Number of slides created initially")


class AddSlideResponse(BaseModel):
    """Response from add_slide"""
    success: bool = Field(description="Whether the operation succeeded")
    slide_id: str = Field(description="The ID of the created slide")
    presentation_id: str = Field(description="The ID of the presentation")
    layout: str = Field(description="The layout used for the slide")


class AddTitleResponse(BaseModel):
    """Response from add_title"""
    success: bool = Field(description="Whether the operation succeeded")
    title_object_id: str = Field(description="The ID of the created title element")
    slide_id: str = Field(description="The ID of the slide containing the title")
    presentation_id: str = Field(description="The ID of the presentation")


class AddBodyTextResponse(BaseModel):
    """Response from add_body_text"""
    success: bool = Field(description="Whether the operation succeeded")
    body_object_id: str = Field(description="The ID of the created body text element")
    slide_id: str = Field(description="The ID of the slide containing the body text")
    presentation_id: str = Field(description="The ID of the presentation")


class AddBodyImageResponse(BaseModel):
    """Response from add_body_image"""
    success: bool = Field(description="Whether the operation succeeded")
    image_object_id: str = Field(description="The ID of the created image element")
    slide_id: str = Field(description="The ID of the slide containing the image")
    presentation_id: str = Field(description="The ID of the presentation")


class AddPageWithContentResponse(BaseModel):
    """Response from add_page_with_content"""
    success: bool = Field(description="Whether the operation succeeded")
    slide_id: str = Field(description="The ID of the created slide")
    presentation_id: str = Field(description="The ID of the presentation")
    layout: str = Field(description="The layout used for the slide")
    elements_added: list[dict[str, str]] = Field(
        description="List of elements added with their types and IDs",
        default_factory=list
    )


class MarkdownToSlidesResponse(BaseModel):
    """Response from create_presentation_from_markdown"""
    success: bool = Field(description="Whether the operation succeeded")
    presentation_id: str = Field(description="The ID of the created presentation")
    presentation_url: str = Field(description="The URL to access the presentation")
    presentation_title: str = Field(description="The title of the presentation")
    slides_created: int = Field(description="Number of slides created")
    total_elements: int = Field(description="Total number of elements created")
    warnings: list[str] = Field(
        description="List of warnings during conversion",
        default_factory=list
    )


class ErrorResponse(BaseModel):
    """Error response for any operation"""
    success: bool = Field(default=False, description="Always False for errors")
    error: str = Field(description="Error message describing what went wrong")
