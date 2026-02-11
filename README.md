# OpenClaw DOCX Parser Plugin

A plugin for [OpenClaw](https://docs.openclaw.ai/) that enables reading and parsing Microsoft Word (.docx) files.

## Features

- Extract content from .docx files in multiple formats:
  - Plain text
  - Markdown (recommended)
  - HTML
- File validation before reading
- Configurable file size limits
- Preserves document structure (headings, lists, links, etc.)
- Comprehensive error handling
- Unit tested with Vitest

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

The plugin registers two tools that AI agents can use:

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
  // Handle warnings if any
  if (result.warnings) {
    console.warn(result.warnings);
  }
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
```

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
├── index.test.ts           # Unit tests
├── skills/
│   └── docx-parser/
│       └── SKILL.md        # Agent guidance
└── README.md
```

## How it Works

The plugin uses [mammoth.js](https://github.com/mwilliamson/mammoth.js) to extract content from .docx files:

1. **Validation**: Checks file extension and size
2. **Extraction**: Uses mammoth to convert .docx to HTML or plain text
3. **Conversion**: Converts HTML to Markdown (if requested)
4. **Return**: Returns content with warnings (if any)

## Limitations

- Only .docx format is supported (not legacy .doc files)
- Maximum file size: 100 MB (configurable)
- Embedded images are not extracted
- Complex formatting may not be perfectly preserved
- Tables are converted to basic text/markdown

## License

MIT

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
