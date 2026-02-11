# DOCX Parser

This plugin provides tools for reading and parsing Microsoft Word (.docx) files.

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

## Usage Guidelines

1. **Always validate first**: Use `docx_validate` before `docx_read` to check file size and format
2. **Choose the right format**:
   - Use `"text"` for simple text extraction
   - Use `"markdown"` for best readability with formatting preserved
   - Use `"html"` if you need full HTML structure
3. **Handle errors**: Always check the `success` field and handle the `error` message
4. **Check warnings**: The `warnings` array may contain messages about content conversion issues

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

- Maximum file size: 100 MB (configurable, default 10 MB)
- Only .docx format is supported (not .doc)
- Complex formatting may not be perfectly preserved
- Embedded images are not extracted
- Tables are converted to basic text/markdown
