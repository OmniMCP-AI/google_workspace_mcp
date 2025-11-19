"""
Markdown to Google Slides Parser

Parses Markdown content into structured slide data for Google Slides creation.
"""

import logging
from typing import List, Optional, Dict, Any
from dataclasses import dataclass
from markdown_it import MarkdownIt
from markdown_it.token import Token

logger = logging.getLogger(__name__)


@dataclass
class SlideData:
    """Represents data for a single slide"""
    title: str
    body_text: str = ""
    image_url: Optional[str] = None
    layout: str = "BLANK"

    def has_content(self) -> bool:
        """Check if slide has any content"""
        return bool(self.title or self.body_text or self.image_url)


class MarkdownToSlidesParser:
    """Parser for converting Markdown to slide structure"""

    def __init__(self):
        self.md = MarkdownIt()
        self.slides: List[SlideData] = []
        self.current_slide: Optional[SlideData] = None
        self.presentation_title: Optional[str] = None
        self.warnings: List[str] = []

    def parse(self, markdown_content: str) -> Dict[str, Any]:
        """
        Parse markdown content into slide structure

        Args:
            markdown_content: Markdown text content

        Returns:
            Dict with presentation_title, slides, and warnings
        """
        tokens = self.md.parse(markdown_content)
        self._process_tokens(tokens)

        # Finalize last slide
        if self.current_slide and self.current_slide.has_content():
            self.slides.append(self.current_slide)

        return {
            "presentation_title": self.presentation_title,
            "slides": self.slides,
            "warnings": self.warnings
        }

    def _process_tokens(self, tokens: List[Token]):
        """Process markdown tokens into slides"""
        i = 0
        while i < len(tokens):
            token = tokens[i]

            if token.type == "heading_open":
                i = self._process_heading(tokens, i)
            elif token.type == "paragraph_open":
                i = self._process_paragraph(tokens, i)
            elif token.type == "bullet_list_open":
                i = self._process_list(tokens, i, ordered=False)
            elif token.type == "ordered_list_open":
                i = self._process_list(tokens, i, ordered=True)
            elif token.type == "fence":
                self._process_code_block(token)
                i += 1
            elif token.type == "image":
                self._process_image(token)
                i += 1
            else:
                i += 1

        return i

    def _process_heading(self, tokens: List[Token], start_idx: int) -> int:
        """Process heading token"""
        heading_open = tokens[start_idx]
        level = int(heading_open.tag[1])  # h1 -> 1, h2 -> 2

        # Find content
        content = ""
        idx = start_idx + 1
        while idx < len(tokens) and tokens[idx].type != "heading_close":
            if tokens[idx].type == "inline":
                content = tokens[idx].content
            idx += 1

        if level == 1:
            # H1 = Presentation title
            if not self.presentation_title:
                self.presentation_title = content
                logger.info(f"Found presentation title: {content}")
            else:
                self.warnings.append(f"Multiple H1 headings found, ignoring: {content}")

        elif level == 2:
            # H2 = New slide
            if self.current_slide and self.current_slide.has_content():
                self.slides.append(self.current_slide)

            self.current_slide = SlideData(title=content)
            logger.info(f"New slide: {content}")

        elif level == 3:
            # H3 = Subheading in body (bold)
            if self.current_slide:
                self._append_to_body(f"\n**{content}**\n")
            else:
                self.warnings.append(f"H3 '{content}' found before any H2, creating new slide")
                self.current_slide = SlideData(title=content)

        else:
            # H4+ = Regular text
            if self.current_slide:
                self._append_to_body(f"\n{content}\n")

        return idx + 1

    def _process_paragraph(self, tokens: List[Token], start_idx: int) -> int:
        """Process paragraph token"""
        idx = start_idx + 1
        content = ""

        while idx < len(tokens) and tokens[idx].type != "paragraph_close":
            if tokens[idx].type == "inline":
                content = self._process_inline(tokens[idx])
            idx += 1

        if content and self.current_slide:
            self._append_to_body(content + "\n")

        return idx + 1

    def _process_inline(self, token: Token) -> str:
        """Process inline content (text with formatting)"""
        content = token.content

        # Process children for images and formatting
        if token.children:
            parts = []
            for child in token.children:
                if child.type == "image":
                    self._process_image(child)
                    # Don't include image markdown in text
                    continue
                elif child.type == "text":
                    parts.append(child.content)
                elif child.type == "code_inline":
                    parts.append(f"`{child.content}`")
                else:
                    parts.append(child.content)
            content = "".join(parts)

        return content

    def _process_list(self, tokens: List[Token], start_idx: int, ordered: bool) -> int:
        """Process bullet or ordered list"""
        idx = start_idx + 1
        list_items = []
        item_number = 1

        while idx < len(tokens):
            token = tokens[idx]

            if token.type in ("bullet_list_close", "ordered_list_close"):
                break

            if token.type == "list_item_open":
                # Get item content
                idx += 1
                item_content = ""
                while idx < len(tokens) and tokens[idx].type != "list_item_close":
                    if tokens[idx].type == "inline":
                        item_content = tokens[idx].content
                    elif tokens[idx].type == "paragraph_open":
                        idx += 1
                        if tokens[idx].type == "inline":
                            item_content = tokens[idx].content
                    idx += 1

                if item_content:
                    if ordered:
                        list_items.append(f"{item_number}. {item_content}")
                        item_number += 1
                    else:
                        list_items.append(f"• {item_content}")

            idx += 1

        if list_items and self.current_slide:
            self._append_to_body("\n".join(list_items) + "\n")

        return idx + 1

    def _process_code_block(self, token: Token):
        """Process code block as monospace text"""
        if self.current_slide:
            code = token.content.strip()
            # Add code with clear boundaries
            self._append_to_body(f"\n```\n{code}\n```\n")

    def _process_image(self, token: Token):
        """Process image - only first image per slide"""
        if not self.current_slide:
            self.warnings.append(f"Image found before any slide: {token.attrGet('src')}")
            return

        image_url = token.attrGet("src")

        if self.current_slide.image_url:
            # Already has an image
            self.warnings.append(
                f"Slide '{self.current_slide.title}': Multiple images found, "
                f"using first one ({self.current_slide.image_url})"
            )
        else:
            self.current_slide.image_url = image_url
            logger.info(f"Added image to slide '{self.current_slide.title}': {image_url}")

    def _append_to_body(self, text: str):
        """Append text to current slide body"""
        if self.current_slide:
            if self.current_slide.body_text:
                self.current_slide.body_text += "\n" + text
            else:
                self.current_slide.body_text = text


def parse_markdown_to_slides(markdown_content: str) -> Dict[str, Any]:
    """
    Parse markdown content into slide structure

    Args:
        markdown_content: Markdown text content

    Returns:
        Dict with:
            - presentation_title: str
            - slides: List[SlideData]
            - warnings: List[str]
    """
    parser = MarkdownToSlidesParser()
    return parser.parse(markdown_content)
