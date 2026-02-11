# OpenClaw DOCX Parser Plugin

A plugin for [OpenClaw](https://docs.openclaw.ai/) that enables reading, writing, and validating Microsoft Word (.docx) files.

## Features

### Reading (.docx files)
- Extract content in multiple formats:
  - Plain text
  - Markdown (recommended)
  - HTML
- File validation before reading
- Configurable file size limits
- Preserves document structure (headings, lists, links, etc.)
- Comprehensive error handling

### Writing (.docx files)
- Create .docx files from Markdown or plain text
- Support for headings, bold, italic, lists, blockquotes, code, links, and more
- Automatic directory creation
- Configurable overwrite behavior
- Document metadata support (title, author)

### Testing
- 41 comprehensive unit tests
- 100% test coverage for core functionality
- Tested with Vitest

## Installation

### Install from local path

```bash
cd /Users/fatelei/github
openclaw plugins install openclaw-docx-parser
```

### Enable the plugin

```bash
openclaw plugins enable docx-parser
openclaw gateway restart
```

## Configuration

Add to your `~/.openclaw/openclaw.json`:

```json
{
  "plugins": {
    "entries": {
      "docx-parser": {
        "enabled": true,
        "config": {
          "maxFileSize": 10,
          "defaultOutputFormat": "markdown"
        }
      }
    }
  }
}
```

### Configuration Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `maxFileSize` | number | 10 | Maximum file size in MB (1-100) |
| `defaultOutputFormat` | string | "markdown" | Default output format: "text", "markdown", or "html" |

## Usage

### Agent Tools

The plugin registers three tools that AI agents can use:

#### `docx_read`

Read and extract content from a .docx file.

```javascript
const result = await docx_read({
  filePath: "/path/to/document.docx",
  outputFormat: "markdown",  // "text" | "markdown" | "html"
  maxFileSize: 10
});

if (result.success) {
  console.log(result.content);
} else {
  console.error(result.error);
}
```

#### `docx_write`

Create a new .docx file from Markdown or text content.

```javascript
const result = await docx_write({
  filePath: "/path/to/output.docx",
  content: "# My Document\n\nThis is **bold** and *italic* text.",
  title: "My Document",
  author: "Author Name",
  overwrite: false
});

if (result.success) {
  console.log(`Created: ${result.filePath} (${result.size} bytes)`);
} else {
  console.error(result.error);
}
```

#### `docx_validate`

Validate a .docx file before reading.

```javascript
const validation = await docx_validate({
  filePath: "/path/to/document.docx",
  maxFileSize: 10
});

if (validation.valid) {
  console.log("File is valid");
  console.log("Size:", validation.size, "bytes");
} else {
  console.error("Validation failed:", validation.error);
}
```

### CLI Commands

```bash
# Read a DOCX file
openclaw docx-read document.docx

# Specify output format
openclaw docx-read document.docx --format html

# Set max file size
openclaw docx-read document.docx --max-size 20

# Validate a file
openclaw docx-validate document.docx

# Write a new DOCX file from markdown
openclaw docx-write output.docx "# Title\n\nContent here"

# Write with title and author
openclaw docx-write report.docx -t "Monthly Report" -a "Jane Doe" -c "# Report\n\nContent"

# Overwrite existing file
openclaw docx-write existing.docx --overwrite "New content"
```

## Markdown Support for Writing

The `docx_write` tool supports the following markdown syntax:

| Element | Syntax | Example |
|---------|--------|---------|
| Heading 1-6 | `#` to `######` | `## Heading` |
| Bold | `**text**` | `**bold text**` |
| Italic | `*text*` | `*italic text*` |
| Inline code | `` `code` `` | `` `code` `` |
| Bulleted list | `- item` | `- Item 1` |
| Numbered list | `1. item` | `1. First item` |
| Blockquote | `> quote` | `> Quote text` |
| Horizontal rule | `---` | `---` |
| Links | `[text](url)` | `[Link](https://example.com)` |

## Development

### Install dependencies

```bash
cd openclaw-docx-parser
npm install
```

### Run tests

```bash
# Run all tests
npm test

# Run tests with UI
npm run test:ui

# Run tests with coverage
npm run test:coverage
```

### Type checking

```bash
npm run typecheck
```

## File Structure

```
openclaw-docx-parser/
├── index.ts                 # Plugin main file
├── openclaw.plugin.json     # Plugin manifest
├── package.json             # Dependencies
├── tsconfig.json           # TypeScript config
├── vitest.config.ts        # Test config
├── index.test.ts           # Unit tests (41 tests)
├── skills/
│   └── docx-parser/
│       └── SKILL.md        # Agent guidance
└── README.md
```

## How it Works

### Reading
The plugin uses [mammoth.js](https://github.com/mwilliamson/mammoth.js) to extract content from .docx files:
1. **Validation**: Checks file extension and size
2. **Extraction**: Uses mammoth to convert .docx to HTML or plain text
3. **Conversion**: Converts HTML to Markdown (if requested)
4. **Return**: Returns content with warnings (if any)

### Writing
The plugin uses [docx](https://docx.js.org/) to create .docx files:
1. **Parsing**: Parses Markdown/text content into structured elements
2. **Document Creation**: Creates a DOCX document with proper formatting
3. **Serialization**: Converts document to buffer and writes to file
4. **Confirmation**: Returns file path and size

## Limitations

### Reading
- Only .docx format is supported (not legacy .doc files)
- Maximum file size: 100 MB (configurable, default 10 MB)
- Embedded images are not extracted
- Complex formatting may not be perfectly preserved
- Tables are converted to basic text/markdown

### Writing
- No image embedding support
- Tables are not supported
- Complex nested formatting may not render perfectly
- Some markdown extensions (footnotes, task lists, etc.) are not supported
- Links are styled but not clickable in all Word viewers

## License

MIT

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
