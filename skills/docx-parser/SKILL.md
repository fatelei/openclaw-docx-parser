# DOCX Parser

This plugin provides tools for reading, writing, and validating Microsoft Word (.docx) files.

## Available Tools

### `docx_read`

Read and extract content from a .docx file.

**Parameters:**
- `filePath` (required): Absolute path to the .docx file
- `outputFormat` (optional): Output format - "text", "markdown", or "html" (default: "markdown")
- `maxFileSize` (optional): Maximum file size in MB (default: 10)

**Returns:**
- `success`: Boolean indicating if the operation succeeded
- `content`: Extracted text/markdown/html content
- `warnings`: Array of warning messages (if any)
- `error`: Error message (if failed)

**Example:**
```javascript
const result = await docx_read({
  filePath: "/path/to/document.docx",
  outputFormat: "markdown"
});
```

### `docx_write`

Create a new .docx file from markdown or text content.

**Parameters:**
- `filePath` (required): Absolute path to the output .docx file
- `content` (required): Content to write (supports markdown or plain text)
- `inputFormat` (optional): Input format - "text" or "markdown" (default: "markdown")
- `title` (optional): Document title
- `author` (optional): Author name
- `overwrite` (optional): Whether to overwrite existing files (default: false)

**Returns:**
- `success`: Boolean indicating if the operation succeeded
- `filePath`: Path to the created file
- `size`: File size in bytes
- `error`: Error message (if failed)

**Example:**
```javascript
const result = await docx_write({
  filePath: "/path/to/output.docx",
  content: "# My Document\n\nThis is **bold** and *italic* text.",
  title: "My Document",
  author: "Author Name"
});
```

### `docx_validate`

Validate a .docx file before reading.

**Parameters:**
- `filePath` (required): Absolute path to the .docx file
- `maxFileSize` (optional): Maximum file size in MB (default: 10)

**Returns:**
- `success`: Boolean indicating if validation succeeded
- `valid`: Boolean indicating if the file is valid
- `size`: File size in bytes (if valid)
- `error`: Error message (if failed)

**Example:**
```javascript
const validation = await docx_validate({
  filePath: "/path/to/document.docx",
  maxFileSize: 10
});
```

## Markdown Support for Writing

The `docx_write` tool supports the following markdown syntax:

| Element | Syntax | Output |
|---------|--------|--------|
| Heading 1 | `# Text` | Large bold heading |
| Heading 2 | `## Text` | Medium bold heading |
| Heading 3-6 | `###` to `######` | Smaller headings |
| Bold | `**text**` or `__text__` | Bold text |
| Italic | `*text*` | Italic text |
| Inline code | `` `code` `` | Monospace red text |
| Bulleted list | `- item` or `* item` | Bulleted list |
| Numbered list | `1. item` | Numbered list |
| Blockquote | `> quote` | Italic gray quote with border |
| Horizontal rule | `---` or `***` | Horizontal line |
| Links | `[text](url)` | Blue underlined link |

## Usage Guidelines

### Reading
1. **Always validate first**: Use `docx_validate` before `docx_read` to check file size and format
2. **Choose the right format**:
   - Use `"text"` for simple text extraction
   - Use `"markdown"` for best readability with formatting preserved
   - Use `"html"` if you need full HTML structure
3. **Handle errors**: Always check the `success` field and handle the `error` message
4. **Check warnings**: The `warnings` array may contain messages about content conversion issues

### Writing
1. **Check file existence**: The tool fails if file exists and `overwrite` is false
2. **Use markdown**: Markdown format provides better structure and formatting
3. **Provide metadata**: Include `title` and `author` for better document properties
4. **Directories**: The tool automatically creates parent directories if needed

## Common Workflows

### Read a document with validation
```javascript
// First validate
const validation = await docx_validate({
  filePath: "/path/to/document.docx"
});

if (!validation.valid) {
  return { error: validation.error };
}

// Then read
const result = await docx_read({
  filePath: "/path/to/document.docx",
  outputFormat: "markdown"
});
```

### Create a new document
```javascript
const result = await docx_write({
  filePath: "/path/to/output.docx",
  content: `# Report Title

## Introduction
This is an introduction with **bold** and *italic* text.

## Key Points
- Point 1
- Point 2
- Point 3

## Conclusion
> This is an important conclusion.
`,
  title: "Monthly Report",
  author: "Jane Doe"
});
```

### Convert document to markdown and back
```javascript
// Read original
const original = await docx_read({
  filePath: "/path/to/input.docx",
  outputFormat: "markdown"
});

// Modify content
const modified = original.content + "\n\n## Additional Notes\nAdded content.";

// Write new file
await docx_write({
  filePath: "/path/to/output.docx",
  content: modified,
  title: "Modified Document"
});
```

### Extract plain text from multiple files
```javascript
const files = ["/path/a.docx", "/path/b.docx"];
const contents = [];

for (const file of files) {
  const result = await docx_read({
    filePath: file,
    outputFormat: "text"
  });
  if (result.success) {
    contents.push(result.content);
  }
}
```

## Limitations

### Reading
- Maximum file size: 100 MB (configurable, default 10 MB)
- Only .docx format is supported (not .doc)
- Complex formatting may not be perfectly preserved
- Embedded images are not extracted
- Tables are converted to basic text/markdown

### Writing
- No image embedding support
- Tables are not supported
- Complex nested formatting may not render perfectly
- Some markdown extensions (footnotes, task lists, etc.) are not supported
- Links are styled but not clickable in all Word viewers
