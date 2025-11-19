# Test Assets

This directory contains test files for the markdown to Google Docs conversion functionality.

## Files

### Markdown Files

- **sample.md** - Basic markdown test file with text formatting, lists, code blocks, and links
- **sample_with_images.md** - Markdown file with embedded local images

### Image Files

- **test_image_1.png** - Blue test image (400x300)
- **test_image_2.png** - Red test image (200x200)

## Usage

These files are used by the test scripts:

- `tests/simple_markdown_test.py` - Tests conversion of `sample.md`
- `tests/test_with_images.py` - Tests conversion of `sample_with_images.md` with local images
- `tests/test_markdown_to_docx.py` - Full integration tests that upload to Google Drive

## Adding Your Own Test Files

You can add your own markdown files to this directory and reference them in tests. When using local images:

1. Place image files in this `assets/` directory
2. Reference them in markdown using relative paths (e.g., `![Description](image.png)`)
3. Make sure to change the working directory to `assets/` before conversion so relative paths resolve correctly

Example:
```python
os.chdir(Path(__file__).parent / "assets")
docx_bytes, temp_path = await convert_markdown_to_docx(markdown_content, ...)
```
