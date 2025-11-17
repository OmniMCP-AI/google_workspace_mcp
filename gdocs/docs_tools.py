"""
Google Docs MCP Tools

This module provides MCP tools for interacting with Google Docs API and managing Google Docs via Drive.
"""
import logging
import asyncio
import io
import os
import re
import tempfile
from typing import List, Optional, Tuple
from uuid import uuid4

from mcp import types
from fastmcp import Context
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# Auth & server utilities
from auth.service_decorator import require_google_service, require_multiple_services
from core.utils import extract_office_xml_text, handle_http_errors
from core.server import server
from core.comments import create_comment_tools

logger = logging.getLogger(__name__)

async def convert_markdown_to_docx(
    markdown_content: str,
    output_path: Optional[str] = None,
    proxy_config: Optional[dict] = None
) -> Tuple[bytes, str]:
    """
    Convert markdown content to docx format, including downloading and embedding images.

    Args:
        markdown_content: The markdown text to convert
        output_path: Optional output path for the docx file. If None, uses a temp file.
        proxy_config: Optional proxy configuration dict with 'http' and 'https' keys

    Returns:
        Tuple of (docx_bytes, output_path)

    Raises:
        ImportError: If pypandoc is not installed
        Exception: If conversion or image download fails
    """
    try:
        import pypandoc
        import requests
    except ImportError as e:
        raise ImportError(
            "pypandoc and requests are required for markdown to docx conversion. "
            "Install with: pip install pypandoc requests python-docx"
        ) from e

    # Use temp directory for downloaded images
    temp_dir = tempfile.mkdtemp()
    downloaded_images = {}

    try:
        # Extract image URLs from markdown
        image_pattern = r'!\[([^\]]*)\]\(([^\)]+)\)'
        images = re.findall(image_pattern, markdown_content)

        # Download images
        for idx, (alt_text, url) in enumerate(images, 1):
            # Skip if already a local file
            if not url.startswith(('http://', 'https://')):
                logger.info(f"[convert_markdown_to_docx] Image {idx} is local path: {url}")
                continue

            try:
                logger.info(f"[convert_markdown_to_docx] Downloading image {idx}: {url}")

                # Prepare request with optional proxy
                request_kwargs = {'timeout': 60}
                if proxy_config:
                    request_kwargs['proxies'] = proxy_config

                response = await asyncio.to_thread(
                    requests.get, url, **request_kwargs
                )

                if response.status_code == 200:
                    # Determine file extension from content-type or URL
                    content_type = response.headers.get('content-type', '')
                    if 'png' in content_type or url.lower().endswith('.png'):
                        ext = 'png'
                    elif 'jpeg' in content_type or 'jpg' in content_type or url.lower().endswith(('.jpg', '.jpeg')):
                        ext = 'jpg'
                    elif 'gif' in content_type or url.lower().endswith('.gif'):
                        ext = 'gif'
                    else:
                        ext = 'png'  # default

                    img_filename = os.path.join(temp_dir, f"img_{idx}.{ext}")
                    with open(img_filename, 'wb') as f:
                        f.write(response.content)
                    downloaded_images[url] = img_filename
                    logger.info(f"[convert_markdown_to_docx] Downloaded {img_filename}")
                else:
                    logger.warning(
                        f"[convert_markdown_to_docx] Failed to download image {idx}: "
                        f"HTTP {response.status_code}"
                    )
            except Exception as e:
                logger.warning(f"[convert_markdown_to_docx] Error downloading image {idx}: {e}")

        # Replace image URLs with local paths in markdown
        processed_markdown = markdown_content
        for url, local_path in downloaded_images.items():
            processed_markdown = processed_markdown.replace(url, local_path)

        # Set output path
        if output_path is None:
            output_path = os.path.join(temp_dir, f"converted_{uuid4().hex}.docx")

        # Convert markdown to docx using pypandoc
        await asyncio.to_thread(
            pypandoc.convert_text,
            processed_markdown,
            'docx',
            format='md',
            outputfile=output_path,
            extra_args=['--resource-path=' + temp_dir]
        )

        # Read the docx file into bytes
        with open(output_path, 'rb') as f:
            docx_bytes = f.read()

        logger.info(f"[convert_markdown_to_docx] Conversion complete: {len(docx_bytes)} bytes")

        return docx_bytes, output_path

    finally:
        # Clean up downloaded images (but keep the output file)
        for img_file in downloaded_images.values():
            try:
                if os.path.exists(img_file):
                    os.remove(img_file)
            except Exception as e:
                logger.warning(f"[convert_markdown_to_docx] Failed to clean up {img_file}: {e}")

def process_tabs_recursively(tabs: List, level: int = 0, target_tab_id: Optional[str] = None) -> List[str]:
    """
    Recursively process tabs and their child tabs.
    
    Args:
        tabs: List of tab objects from Google Docs API
        level: Current nesting level for indentation
        target_tab_id: If specified, only process this specific tab ID
        
    Returns:
        List[str]: List of processed text lines from all tabs (or specific tab)
    """
    processed_lines: List[str] = []
    indent = "  " * level  # Indentation based on nesting level
    
    for i, tab in enumerate(tabs):
        tab_properties = tab.get('tabProperties', {})
        tab_title = tab_properties.get('title', f'Tab {i+1}')
        tab_id = tab_properties.get('tabId', 'unknown')
        
        # If target_tab_id is specified, skip tabs that don't match
        if target_tab_id and tab_id != target_tab_id:
            # Still check child tabs recursively
            child_tabs = tab.get('childTabs', [])
            if child_tabs:
                child_lines = process_tabs_recursively(child_tabs, level + 1, target_tab_id)
                processed_lines.extend(child_lines)
            
            nested_tabs = tab.get('tabs', [])
            if nested_tabs:
                nested_lines = process_tabs_recursively(nested_tabs, level + 1, target_tab_id)
                processed_lines.extend(nested_lines)
            continue
        
        logger.info(f"[process_tabs_recursively] Processing tab at level {level}: '{tab_title}' (ID: {tab_id})")
        processed_lines.append(f"\n{indent}=== TAB {i+1}: {tab_title} (ID: {tab_id}) ===\n")
        
        # Process document content for this tab
        document_tab = tab.get('documentTab', {})
        if document_tab:
            tab_body = document_tab.get('body', {})
            if tab_body:
                tab_content = tab_body.get('content', [])
                logger.info(f"[process_tabs_recursively] Tab {i} has {len(tab_content)} content elements")
                
                if tab_content:
                    tab_processed_content = process_structural_elements(tab_content)
                    # Add indentation to content lines
                    for line in tab_processed_content:
                        processed_lines.append(f"{indent}{line}")
                else:
                    processed_lines.append(f"{indent}[EMPTY TAB CONTENT]\n")
            else:
                processed_lines.append(f"{indent}[NO BODY CONTENT]\n")
        else:
            processed_lines.append(f"{indent}[NO DOCUMENT TAB CONTENT]\n")
        
        # Process child tabs recursively
        child_tabs = tab.get('childTabs', [])
        if child_tabs:
            logger.info(f"[process_tabs_recursively] Tab '{tab_title}' has {len(child_tabs)} child tabs")
            processed_lines.append(f"{indent}--- CHILD TABS ---\n")
            processed_lines.extend(process_tabs_recursively(child_tabs, level + 1, target_tab_id))
            processed_lines.append(f"{indent}--- END CHILD TABS ---\n")
        
        # Also check for nested tabs in different structure
        nested_tabs = tab.get('tabs', [])
        if nested_tabs:
            logger.info(f"[process_tabs_recursively] Tab '{tab_title}' has {len(nested_tabs)} nested tabs")
            processed_lines.append(f"{indent}--- NESTED TABS ---\n")
            processed_lines.extend(process_tabs_recursively(nested_tabs, level + 1, target_tab_id))
            processed_lines.append(f"{indent}--- END NESTED TABS ---\n")
    
    return processed_lines

def process_structural_elements(elements: List) -> List[str]:
    """
    Process various types of structural elements in a Google Doc.
    
    Args:
        elements: List of structural elements from Google Docs API
        
    Returns:
        List[str]: List of processed text lines
    """
    processed_lines: List[str] = []
    
    for element in elements:
        if 'paragraph' in element:
            # Handle paragraph elements
            paragraph = element.get('paragraph', {})
            para_elements = paragraph.get('elements', [])
            current_line_text = ""
            
            for pe in para_elements:
                text_run = pe.get('textRun', {})
                if text_run and 'content' in text_run:
                    current_line_text += text_run['content']
            
            if current_line_text.strip():
                processed_lines.append(current_line_text)
                
        elif 'table' in element:
            # Handle table elements
            table = element.get('table', {})
            processed_lines.append("\n--- TABLE ---\n")
            
            table_rows = table.get('tableRows', [])
            for row in table_rows:
                row_cells = row.get('tableCells', [])
                cell_texts = []
                
                for cell in row_cells:
                    cell_content = cell.get('content', [])
                    cell_text = "".join(process_structural_elements(cell_content))
                    cell_texts.append(cell_text.strip())
                
                if any(cell_texts):  # Only add row if it has content
                    processed_lines.append(" | ".join(cell_texts) + "\n")
            
            processed_lines.append("--- END TABLE ---\n")
            
        elif 'sectionBreak' in element:
            # Handle section breaks
            processed_lines.append("\n--- SECTION BREAK ---\n")
            
        elif 'tableOfContents' in element:
            # Handle table of contents
            processed_lines.append("\n--- TABLE OF CONTENTS ---\n")
            
        elif 'pageBreak' in element:
            # Handle page breaks
            processed_lines.append("\n--- PAGE BREAK ---\n")
            
        elif 'horizontalRule' in element:
            # Handle horizontal rules
            processed_lines.append("\n--- HORIZONTAL RULE ---\n")
            
        elif 'footerContent' in element:
            # Handle footer content
            processed_lines.append("\n--- FOOTER ---\n")
            footer_content = element.get('footerContent', {}).get('content', [])
            processed_lines.extend(process_structural_elements(footer_content))
            processed_lines.append("--- END FOOTER ---\n")
            
        elif 'headerContent' in element:
            # Handle header content
            processed_lines.append("\n--- HEADER ---\n")
            header_content = element.get('headerContent', {}).get('content', [])
            processed_lines.extend(process_structural_elements(header_content))
            processed_lines.append("--- END HEADER ---\n")
            
        # Add more element types as needed
        # 'pageBreak', 'horizontalRule', etc.
    
    return processed_lines

@server.tool
@require_google_service("drive", "drive_read")
@handle_http_errors("search_docs")
async def search_docs(
    service,
    ctx: Context,
    query: str,
    page_size: int = 10,
    user_google_email: Optional[str] = None,
):
    """
    <description>Searches for Google Docs by document name using Drive API with Google Docs MIME type filtering. Returns document metadata including IDs, names, modification times, and web view links for up to 10 documents.</description>
    
    <use_case>Finding specific Google Docs for content processing, locating documents by partial name matches, or discovering recently modified docs for collaborative editing workflows.</use_case>
    
    <limitation>Searches only document titles, not content. Limited to Google Docs format only - excludes Word files or other document formats. Cannot search within document text or comments.</limitation>
    
    <failure_cases>Fails with malformed search queries containing special characters, when user lacks Drive access permissions, or if Google Drive API quotas are exceeded.</failure_cases>

    Returns:
        str: A formatted list of Google Docs matching the search query.
    """
    logger.info(f"[search_docs] Email={user_google_email}, Query='{query}'")

    escaped_query = query.replace("'", "\\'")

    response = await asyncio.to_thread(
        service.files().list(
            q=f"name contains '{escaped_query}' and mimeType='application/vnd.google-apps.document' and trashed=false",
            pageSize=page_size,
            fields="files(id, name, createdTime, modifiedTime, webViewLink)"
        ).execute
    )
    files = response.get('files', [])
    if not files:
        return f"No Google Docs found matching '{query}'."

    output = [f"Found {len(files)} Google Docs matching '{query}':"]
    for f in files:
        output.append(
            f"- {f['name']} (ID: {f['id']}) Modified: {f.get('modifiedTime')} Link: {f.get('webViewLink')}"
        )
    return "\n".join(output)

@server.tool
@require_multiple_services([
    {"service_type": "drive", "scopes": "drive_read", "param_name": "drive_service"},
    {"service_type": "docs", "scopes": "docs_read", "param_name": "docs_service"}
])
@handle_http_errors("get_doc_content")
async def get_doc_content(
    ctx: Context,
    drive_service: str,
    docs_service: str,
    document_id: str,
    user_google_email: Optional[str] = None,
    tab_id: Optional[str] = None,
):
    """
    <description>Extracts full text content from Google Docs (including tabbed documents) and Office files (.docx) stored in Drive. Processes complex document structures, tables, headers/footers, and multiple tabs into readable plain text.</description>
    
    <use_case>Extracting document content for analysis, processing multi-tab Google Docs for content migration, or converting Office documents to text for automated workflows.</use_case>
    
    <limitation>Returns plain text only - formatting, images, and complex layouts are lost. Limited to documents under 50MB. Cannot process password-protected or heavily corrupted files.</limitation>
    
    <failure_cases>Fails with invalid document IDs, documents the user cannot access due to permissions, corrupted Office files, or when specifying invalid tab IDs for tabbed documents.</failure_cases>
    
    Args:
        document_id: The ID of the document to retrieve
        user_google_email: Optional user email for context
        tab_id: Optional tab ID to retrieve content from a specific tab only

    Returns:
        str: The document content with metadata header.
    """
    logger.info(f"[get_doc_content] Invoked. Document/File ID: '{document_id}' for user '{user_google_email}', tab_id: '{tab_id}'")

    # Step 2: Get file metadata from Drive
    file_metadata = await asyncio.to_thread(
        drive_service.files().get(
            fileId=document_id, fields="id, name, mimeType, webViewLink"
        ).execute
    )
    mime_type = file_metadata.get("mimeType", "")
    file_name = file_metadata.get("name", "Unknown File")
    web_view_link = file_metadata.get("webViewLink", "#")

    logger.info(f"[get_doc_content] File '{file_name}' (ID: {document_id}) has mimeType: '{mime_type}'")

    body_text = "" # Initialize body_text

    # Step 3: Process based on mimeType
    if mime_type == "application/vnd.google-apps.document":
        logger.info(f"[get_doc_content] Processing as native Google Doc.")
        doc_data = await asyncio.to_thread(
            docs_service.documents().get(
                documentId=document_id,
                includeTabsContent=True
            ).execute
        )
        logger.info(f"[get_doc_content] Processing as native Google Doc.")
        
        # Process tabs if they exist
        tabs = doc_data.get('tabs', [])
        logger.info(f"[get_doc_content] Found {len(tabs)} tabs")
        
        # Debug: Print full document structure
        logger.info(f"[get_doc_content] Document keys: {list(doc_data.keys())}")
        if tabs:
            for i, tab in enumerate(tabs):
                logger.info(f"[get_doc_content] Tab {i} keys: {list(tab.keys())}")
                logger.info(f"[get_doc_content] Tab {i} properties: {tab.get('tabProperties', {})}")
                
                # Check for child tabs in all possible locations
                child_tabs = tab.get('childTabs', [])
                nested_tabs = tab.get('tabs', [])
                document_tab = tab.get('documentTab', {})
                
                logger.info(f"[get_doc_content] Tab {i} has {len(child_tabs)} childTabs")
                logger.info(f"[get_doc_content] Tab {i} has {len(nested_tabs)} nested tabs")
                logger.info(f"[get_doc_content] Tab {i} documentTab keys: {list(document_tab.keys()) if document_tab else 'None'}")
                
                if document_tab:
                    tab_body = document_tab.get('body', {})
                    if tab_body:
                        tab_content = tab_body.get('content', [])
                        logger.info(f"[get_doc_content] Tab {i} body content elements: {len(tab_content)}")
                        # Log first few elements to understand structure
                        for j, element in enumerate(tab_content[:3]):
                            logger.info(f"[get_doc_content] Tab {i} element {j}: {list(element.keys())}")
                    else:
                        logger.info(f"[get_doc_content] Tab {i} has no body")
        
        processed_text_lines: List[str] = []
        
        if tabs:
            # Document has tabs - process all tabs recursively or filter by tab_id
            if tab_id:
                logger.info(f"[get_doc_content] Processing specific tab_id '{tab_id}' from {len(tabs)} tabs")
                processed_text_lines.extend(process_tabs_recursively(tabs, 0, tab_id))
                # Check if we found any content for the specified tab_id
                if not processed_text_lines:
                    logger.warning(f"[get_doc_content] No content found for tab_id '{tab_id}'")
                    return f'File: "{file_name}" (ID: {document_id}, Type: {mime_type})\nLink: {web_view_link}\n\n--- ERROR ---\nNo tab found with tab_id "{tab_id}".'
            else:
                logger.info(f"[get_doc_content] Processing all {len(tabs)} tabs recursively")
                processed_text_lines.extend(process_tabs_recursively(tabs, 0))
        else:
            # Document without tabs - process body content directly
            if tab_id:
                logger.warning(f"[get_doc_content] tab_id '{tab_id}' specified but document has no tabs")
                return f'File: "{file_name}" (ID: {document_id}, Type: {mime_type})\nLink: {web_view_link}\n\n--- ERROR ---\nSpecified tab_id "{tab_id}" but document has no tabs.'
            body_elements = doc_data.get('body', {}).get('content', [])
            processed_text_lines.extend(process_structural_elements(body_elements))
            
        body_text = "".join(processed_text_lines)
    else:
        logger.info(f"[get_doc_content] Processing as Drive file (e.g., .docx, other). MimeType: {mime_type}")
        
        # tab_id is not supported for non-Google Docs files
        if tab_id:
            logger.warning(f"[get_doc_content] tab_id '{tab_id}' specified but not supported for non-Google Docs files")
            return f'File: "{file_name}" (ID: {document_id}, Type: {mime_type})\nLink: {web_view_link}\n\n--- ERROR ---\ntab_id is not supported for non-Google Docs files.'

        export_mime_type_map = {
                # Example: "application/vnd.google-apps.spreadsheet"z: "text/csv",
                # Native GSuite types that are not Docs would go here if this function
                # was intended to export them. For .docx, direct download is used.
        }
        effective_export_mime = export_mime_type_map.get(mime_type)

        request_obj = (
            drive_service.files().export_media(fileId=document_id, mimeType=effective_export_mime)
            if effective_export_mime
            else drive_service.files().get_media(fileId=document_id)
        )

        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request_obj)
        loop = asyncio.get_event_loop()
        done = False
        while not done:
            status, done = await loop.run_in_executor(None, downloader.next_chunk)

        file_content_bytes = fh.getvalue()

        office_text = extract_office_xml_text(file_content_bytes, mime_type)
        if office_text:
            body_text = office_text
        else:
            try:
                body_text = file_content_bytes.decode("utf-8")
            except UnicodeDecodeError:
                body_text = (
                    f"[Binary or unsupported text encoding for mimeType '{mime_type}' - "
                    f"{len(file_content_bytes)} bytes]"
                )

    header = (
        f'File: "{file_name}" (ID: {document_id}, Type: {mime_type})\n'
        f'Link: {web_view_link}\n\n--- CONTENT ---\n'
    )
    print(body_text)
    return header + body_text

@server.tool
@require_google_service("drive", "drive_read")
@handle_http_errors("list_docs_in_folder")
async def list_docs_in_folder(
    service,
    ctx: Context,
    user_google_email: Optional[str] = None,
    folder_id: str = 'root',
    page_size: int = 100
):
    """
    <description>Lists all Google Docs within a specific Drive folder showing document names, IDs, modification times, and web view links. Returns up to 100 documents per page with pagination support for large folders.</description>
    
    <use_case>Organizing document workflows by folder, bulk processing documents in specific directories, or discovering documents within project folders for content analysis.</use_case>
    
    <limitation>Limited to Google Docs format only - excludes Word files or other document types. Shows only immediate folder contents, not recursive subfolder documents. Requires valid folder access permissions.</limitation>
    
    <failure_cases>Fails with invalid folder IDs, folders the user cannot access due to sharing restrictions, or when trying to list contents of files instead of folders.</failure_cases>

    Returns:
        str: A formatted list of Google Docs in the specified folder.
    """
    logger.info(f"[list_docs_in_folder] Invoked. Email: '{user_google_email}', Folder ID: '{folder_id}'")

    rsp = await asyncio.to_thread(
        service.files().list(
            q=f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.document' and trashed=false",
            pageSize=page_size,
            fields="files(id, name, modifiedTime, webViewLink)"
        ).execute
    )
    items = rsp.get('files', [])
    if not items:
        return f"No Google Docs found in folder '{folder_id}'."
    out = [f"Found {len(items)} Docs in folder '{folder_id}':"]
    for f in items:
        out.append(f"- {f['name']} (ID: {f['id']}) Modified: {f.get('modifiedTime')} Link: {f.get('webViewLink')}")
    return "\n".join(out)

@server.tool
@require_google_service("docs", "docs_write")
@handle_http_errors("create_doc")
async def create_doc(
    service,
    ctx: Context,
    title: str,
    content: str,
    user_google_email: Optional[str] = None,
):
    """
    <description>Creates a new Google Docs document with specified title and optional initial plain text content. Document is immediately accessible via web interface and ready for collaborative editing.</description>
    
    <use_case>Creating new documents for reports, initializing document templates with starter content, or generating documents programmatically for automated workflows.</use_case>
    
    <limitation>Supports only plain text initial content - no formatting, images, or complex structures. Cannot create documents in specific folders during creation - requires separate sharing/moving operations.</limitation>
    
    <failure_cases>Fails when user lacks Google Docs creation permissions, if title exceeds character limits, or if Google Drive storage quota is exceeded.</failure_cases>

    Returns:
        str: Confirmation message with document ID and link.
    """
    logger.info(f"[create_doc] Invoked. Email: '{user_google_email}', Title='{title}'")

    doc = await asyncio.to_thread(service.documents().create(body={'title': title}).execute)
    doc_id = doc.get('documentId')
    if content:
        requests = [{'insertText': {'location': {'index': 1}, 'text': content}}]
        await asyncio.to_thread(service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute)
    link = f"https://docs.google.com/document/d/{doc_id}/edit"
    msg = f"Created Google Doc '{title}' (ID: {doc_id}) for {user_google_email}. Link: {link}"
    logger.info(f"Successfully created Google Doc '{title}' (ID: {doc_id}) for {user_google_email}. Link: {link}")
    return msg

async def _create_doc_from_markdown_impl(
    drive_service,
    title: str,
    markdown_content: str,
) -> str:
    """
    Internal implementation for creating a Google Doc from markdown content.

    This function is called by both the MCP tool and can be used directly in tests.

    Args:
        drive_service: Google Drive service object
        title: The title for the new Google Doc
        markdown_content: The markdown content to convert and upload

    Returns:
        str: Confirmation message with document ID and link.
    """
    logger.info(f"[create_doc_from_markdown] Invoked. Title='{title}'")

    try:
        # Convert markdown to docx (proxy will be handled by system environment)
        docx_bytes, temp_docx_path = await convert_markdown_to_docx(
            markdown_content=markdown_content,
            output_path=None,  # Use automatic temp file
            proxy_config=None  # Use system proxy settings
        )

        logger.info(f"[create_doc_from_markdown] Converted markdown to docx: {len(docx_bytes)} bytes")

        # Upload docx to Google Drive and convert to Google Docs
        file_metadata = {
            'name': title,
            'mimeType': 'application/vnd.google-apps.document'  # Convert to Google Docs format
        }

        media = io.BytesIO(docx_bytes)

        created_file = await asyncio.to_thread(
            drive_service.files().create(
                body=file_metadata,
                media_body=MediaIoBaseUpload(
                    media,
                    mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    resumable=True
                ),
                fields='id, name, webViewLink',
                supportsAllDrives=True
            ).execute
        )

        doc_id = created_file.get('id')
        web_link = created_file.get('webViewLink')

        # Clean up temp docx file
        try:
            if temp_docx_path and os.path.exists(temp_docx_path):
                os.remove(temp_docx_path)
                # Also try to remove the temp directory if empty
                temp_dir = os.path.dirname(temp_docx_path)
                try:
                    os.rmdir(temp_dir)
                except OSError:
                    pass  # Directory not empty or other issue, ignore
        except Exception as e:
            logger.warning(f"[create_doc_from_markdown] Failed to clean up temp file: {e}")

        msg = f"Created Google Doc from markdown '{title}' (ID: {doc_id}). Link: {web_link}"
        logger.info(msg)
        return msg

    except ImportError as e:
        error_msg = (
            f"Failed to create document from markdown: {str(e)}. "
            "Please ensure pypandoc is installed: pip install pypandoc"
        )
        logger.error(f"[create_doc_from_markdown] {error_msg}")
        return f"Error: {error_msg}"
    except Exception as e:
        error_msg = f"Failed to create document from markdown: {str(e)}"
        logger.error(f"[create_doc_from_markdown] {error_msg}")
        raise

@server.tool
@require_multiple_services([
    {"service_type": "drive", "scopes": "drive_write", "param_name": "drive_service"},
])
@handle_http_errors("create_doc_from_markdown")
async def create_doc_from_markdown(
    ctx: Context,
    drive_service: str,
    title: str,
    markdown_content: str,
    user_google_email: Optional[str] = None,
):
    """
    <description>Creates a new Google Docs document from markdown content by converting it to docx format first. Supports markdown features including headings, bold/italic text, images (downloaded from URLs), links, lists, and code blocks. Uses system proxy settings if configured.</description>

    <use_case>Creating formatted documents from markdown files, converting documentation to Google Docs with proper formatting and embedded images, or migrating markdown-based content to Google Workspace.</use_case>

    <limitation>Requires pypandoc to be installed on the system. Images must be publicly accessible URLs or local files. Complex markdown features may not translate perfectly to Google Docs format. Large images may take time to download.</limitation>

    <failure_cases>Fails if pypandoc is not installed, if image URLs are inaccessible or behind authentication, or if Google Drive storage quota is exceeded.</failure_cases>

    Args:
        title: The title for the new Google Doc
        markdown_content: The markdown content to convert and upload
        user_google_email: Optional user email for logging context

    Returns:
        str: Confirmation message with document ID and link.
    """
    if user_google_email:
        logger.info(f"[create_doc_from_markdown] User: {user_google_email}")

    return await _create_doc_from_markdown_impl(
        drive_service=drive_service,
        title=title,
        markdown_content=markdown_content,
    )


# Create comment management tools for documents
_comment_tools = create_comment_tools("document", "document_id")

# Extract and register the functions
read_doc_comments = _comment_tools['read_comments']
create_doc_comment = _comment_tools['create_comment']
reply_to_comment = _comment_tools['reply_to_comment']
resolve_comment = _comment_tools['resolve_comment']

if __name__ == '__main__':
    asyncio.run(get_doc_content(drive_service="drive", docs_service="docs", document_id="18-52JXU073R9wQ6ip-MrKMjHVq2QgR1VIEFXgMGyuI8"))
