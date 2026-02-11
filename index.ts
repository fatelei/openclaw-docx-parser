/**
 * OpenClaw DOCX Parser Plugin
 *
 * A plugin for reading and writing Microsoft Word (.docx) files.
 * Supports extracting text, markdown, and HTML content from DOCX files.
 * Supports creating DOCX files from markdown or plain text content.
 */

import * as mammoth from 'mammoth';
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, UnderlineType } from 'docx';
import * as fs from 'fs/promises';
import * as path from 'path';

interface DocxReadOptions {
  /** Absolute path to the .docx file */
  filePath: string;
  /** Output format: 'text' for plain text, 'markdown' for markdown, 'html' for HTML */
  outputFormat?: 'text' | 'markdown' | 'html';
  /** Maximum file size in MB (default: 10) */
  maxFileSize?: number;
}

interface DocxReadResult {
  success: boolean;
  content?: string;
  metadata?: {
    title?: string;
    author?: string;
    lastModifiedBy?: string;
    revision?: number;
    createdAt?: string;
    modifiedAt?: string;
  };
  warnings?: string[];
  error?: string;
}

interface DocxValidateOptions {
  /** Absolute path to the .docx file */
  filePath: string;
  /** Maximum file size in MB (default: 10) */
  maxFileSize?: number;
}

interface DocxValidateResult {
  success: boolean;
  valid: boolean;
  size?: number;
  error?: string;
}

interface DocxWriteOptions {
  /** Absolute path to the output .docx file */
  filePath: string;
  /** Content to write (supports markdown or plain text) */
  content: string;
  /** Content format: 'text' or 'markdown' */
  inputFormat?: 'text' | 'markdown';
  /** Document title (optional) */
  title?: string;
  /** Author name (optional) */
  author?: string;
  /** Whether to overwrite existing file (default: false) */
  overwrite?: boolean;
}

interface DocxWriteResult {
  success: boolean;
  filePath?: string;
  size?: number;
  error?: string;
}

/**
 * Convert HTML to Markdown format
 */
function htmlToMarkdown(html: string): string {
  return html
    // Headings
    .replace(/<h1[^>]*>(.*?)<\/h1>/gi, '# $1\n\n')
    .replace(/<h2[^>]*>(.*?)<\/h2>/gi, '## $1\n\n')
    .replace(/<h3[^>]*>(.*?)<\/h3>/gi, '### $1\n\n')
    .replace(/<h4[^>]*>(.*?)<\/h4>/gi, '#### $1\n\n')
    .replace(/<h5[^>]*>(.*?)<\/h5>/gi, '##### $1\n\n')
    .replace(/<h6[^>]*>(.*?)<\/h6>/gi, '###### $1\n\n')
    // Bold and italic
    .replace(/<strong[^>]*>(.*?)<\/strong>/gi, '**$1**')
    .replace(/<b[^>]*>(.*?)<\/b>/gi, '**$1**')
    .replace(/<em[^>]*>(.*?)<\/em>/gi, '*$1*')
    .replace(/<i[^>]*>(.*?)<\/i>/gi, '*$1*')
    // Links
    .replace(/<a[^>]*href="([^"]*)"[^>]*>(.*?)<\/a>/gi, '[$2]($1)')
    // Images
    .replace(/<img[^>]*src="([^"]*)"[^>]*alt="([^"]*)"[^>]*>/gi, '![$2]($1)')
    .replace(/<img[^>]*alt="([^"]*)"[^>]*src="([^"]*)"[^>]*>/gi, '![$1]($2)')
    // Line breaks
    .replace(/<br\s*\/?>/gi, '\n')
    // Paragraphs
    .replace(/<\/p>\s*<p[^>]*>/gi, '\n\n')
    .replace(/<p[^>]*>(.*?)<\/p>/gi, '$1\n\n')
    // Lists (basic support)
    .replace(/<ul[^>]*>/gi, '')
    .replace(/<\/ul>/gi, '\n')
    .replace(/<ol[^>]*>/gi, '')
    .replace(/<\/ol>/gi, '\n')
    .replace(/<li[^>]*>(.*?)<\/li>/gi, '- $1\n')
    // Code blocks
    .replace(/<pre[^>]*><code[^>]*>([\s\S]*?)<\/code><\/pre>/gi, '```\n$1\n```\n')
    .replace(/<code[^>]*>(.*?)<\/code>/gi, '`$1`')
    // Blockquotes
    .replace(/<blockquote[^>]*>(.*?)<\/blockquote>/gi, '> $1\n\n')
    // Horizontal rules
    .replace(/<hr\s*\/?>/gi, '---\n\n')
    // Tables (basic conversion)
    .replace(/<table[^>]*>/gi, '\n')
    .replace(/<\/table>/gi, '\n')
    .replace(/<tr[^>]*>/gi, '')
    .replace(/<\/tr>/gi, '\n')
    .replace(/<t[dh][^>]*>(.*?)<\/t[dh]>/gi, '| $1 ')
    // Remove remaining tags
    .replace(/<[^>]+>/g, '')
    // Clean up extra whitespace
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/**
 * Validate a DOCX file
 */
async function validateDocxFile(
  options: DocxValidateOptions
): Promise<DocxValidateResult> {
  const { filePath, maxFileSize = 10 } = options;

  try {
    // Check file exists
    const stats = await fs.stat(filePath);

    if (!stats.isFile()) {
      return {
        success: true,
        valid: false,
        error: 'Path is not a file',
      };
    }

    // Check file extension
    if (!filePath.toLowerCase().endsWith('.docx')) {
      return {
        success: true,
        valid: false,
        error: 'File must have .docx extension',
      };
    }

    // Check file size
    const maxSizeBytes = maxFileSize * 1024 * 1024;
    if (stats.size > maxSizeBytes) {
      return {
        success: true,
        valid: false,
        size: stats.size,
        error: `File size (${(stats.size / 1024 / 1024).toFixed(2)}MB) exceeds maximum (${maxFileSize}MB)`,
      };
    }

    return {
      success: true,
      valid: true,
      size: stats.size,
    };
  } catch (error: any) {
    return {
      success: false,
      valid: false,
      error: error.message || 'File validation failed',
    };
  }
}

/**
 * Read and parse a DOCX file
 */
async function readDocxFile(options: DocxReadOptions): Promise<DocxReadResult> {
  const { filePath, outputFormat = 'markdown', maxFileSize = 10 } = options;

  // Validate file first
  const validation = await validateDocxFile({ filePath, maxFileSize });

  if (!validation.success) {
    return {
      success: false,
      error: validation.error || 'Validation failed',
    };
  }

  if (!validation.valid) {
    return {
      success: false,
      error: validation.error || 'Invalid file',
    };
  }

  try {
    // Extract content based on output format
    let result: mammoth.Result<any>;
    let content: string;

    if (outputFormat === 'html') {
      result = await mammoth.convertToHtml({ path: filePath });
      content = result.value;
    } else if (outputFormat === 'markdown') {
      // Convert to HTML first, then to markdown
      result = await mammoth.convertToHtml({ path: filePath });
      content = htmlToMarkdown(result.value);
    } else {
      // Plain text
      result = await mammoth.extractRawText({ path: filePath });
      content = result.value;
    }

    // Extract metadata if available
    let metadata: DocxReadResult['metadata'] = undefined;

    return {
      success: true,
      content,
      metadata,
      warnings: result.messages.length > 0 ? result.messages : undefined,
    };
  } catch (error: any) {
    return {
      success: false,
      error: error.message || 'Failed to read DOCX file',
    };
  }
}

/**
 * Parse markdown-style text and convert to DOCX paragraphs
 */
function parseMarkdownToParagraphs(markdown: string, title?: string): Paragraph[] {
  const paragraphs: Paragraph[] = [];
  const lines = markdown.split('\n');

  // Add title if provided
  if (title) {
    paragraphs.push(
      new Paragraph({
        text: title,
        heading: HeadingLevel.TITLE,
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
      })
    );
  }

  for (const line of lines) {
    const trimmed = line.trim();

    // Skip empty lines
    if (!trimmed) {
      paragraphs.push(new Paragraph({ text: '' }));
      continue;
    }

    // Headings
    if (trimmed.startsWith('# ')) {
      paragraphs.push(
        new Paragraph({
          children: [new TextRun({ text: trimmed.slice(2), bold: true, size: 32 })],
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 200, after: 100 },
        })
      );
    } else if (trimmed.startsWith('## ')) {
      paragraphs.push(
        new Paragraph({
          children: [new TextRun({ text: trimmed.slice(3), bold: true, size: 28 })],
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 180, after: 100 },
        })
      );
    } else if (trimmed.startsWith('### ')) {
      paragraphs.push(
        new Paragraph({
          children: [new TextRun({ text: trimmed.slice(4), bold: true, size: 26 })],
          heading: HeadingLevel.HEADING_3,
          spacing: { before: 160, after: 100 },
        })
      );
    } else if (trimmed.startsWith('#### ')) {
      paragraphs.push(
        new Paragraph({
          children: [new TextRun({ text: trimmed.slice(5), bold: true, size: 24 })],
          heading: HeadingLevel.HEADING_4,
          spacing: { before: 140, after: 100 },
        })
      );
    } else if (trimmed.startsWith('##### ')) {
      paragraphs.push(
        new Paragraph({
          children: [new TextRun({ text: trimmed.slice(6), bold: true, size: 22 })],
          heading: HeadingLevel.HEADING_5,
          spacing: { before: 120, after: 100 },
        })
      );
    } else if (trimmed.startsWith('###### ')) {
      paragraphs.push(
        new Paragraph({
          children: [new TextRun({ text: trimmed.slice(7), bold: true, size: 20 })],
          heading: HeadingLevel.HEADING_6,
          spacing: { before: 100, after: 100 },
        })
      );
    }
    // Bold text: **text** or __text__
    else if (trimmed.match(/^\*\*(.+)\*\*$/) || trimmed.match(/^__(.+)__$/)) {
      const match = trimmed.match(/^\*\*(.+)\*\*$/) || trimmed.match(/^__(.+)__$/);
      paragraphs.push(
        new Paragraph({
          children: [new TextRun({ text: match![1], bold: true })],
          spacing: { after: 100 },
        })
      );
    }
    // Italic text: *text* or _text_
    else if (trimmed.match(/^\*(.+)\*$/) && !trimmed.startsWith('**')) {
      const match = trimmed.match(/^\*(.+)\*$/);
      paragraphs.push(
        new Paragraph({
          children: [new TextRun({ text: match![1], italics: true })],
          spacing: { after: 100 },
        })
      );
    }
    // Lists: - item or * item
    else if (trimmed.startsWith('- ') || trimmed.startsWith('* ')) {
      paragraphs.push(
        new Paragraph({
          children: [
            new TextRun({ text: '• ' }),
            new TextRun({ text: trimmed.slice(2) }),
          ],
          indent: { left: 720 },
          spacing: { after: 60 },
        })
      );
    }
    // Numbered lists: 1. item
    else if (trimmed.match(/^\d+\.\s/)) {
      const match = trimmed.match(/^\d+\.\s(.+)$/);
      paragraphs.push(
        new Paragraph({
          children: [
            new TextRun({ text: match![1] }),
          ],
          numbering: {
            reference: 'default-numbering',
            level: 0,
          },
          spacing: { after: 60 },
        })
      );
    }
    // Blockquotes: > text
    else if (trimmed.startsWith('> ')) {
      paragraphs.push(
        new Paragraph({
          children: [
            new TextRun({
              text: trimmed.slice(2),
              italics: true,
              color: '666666',
            }),
          ],
          indent: { left: 720 },
          border: {
            left: {
              color: 'CCCCCC',
              space: 120,
              style: 'single',
              size: 6,
            },
          },
          spacing: { after: 100 },
        })
      );
    }
    // Horizontal rule: --- or ***
    else if (trimmed === '---' || trimmed === '***') {
      paragraphs.push(
        new Paragraph({
          children: [
            new TextRun({
              text: '',
            }),
          ],
          border: {
            bottom: {
              color: 'CCCCCC',
              space: 100,
              style: 'single',
              size: 6,
            },
          },
          spacing: { before: 200, after: 200 },
        })
      );
    }
    // Inline formatting within text
    else {
      // Parse inline formatting
      const textRuns = parseInlineFormatting(trimmed);
      paragraphs.push(
        new Paragraph({
          children: textRuns,
          spacing: { after: 100 },
        })
      );
    }
  }

  return paragraphs;
}

/**
 * Parse inline formatting (bold, italic, code, links)
 */
function parseInlineFormatting(text: string): TextRun[] {
  const runs: TextRun[] = [];
  let current = '';
  let i = 0;

  while (i < text.length) {
    // Bold: **text**
    if (text.slice(i).startsWith('**') && text.slice(i + 2).includes('**')) {
      if (current) {
        runs.push(new TextRun(current));
        current = '';
      }
      const end = text.indexOf('**', i + 2);
      runs.push(new TextRun({ text: text.slice(i + 2, end), bold: true }));
      i = end + 2;
      continue;
    }

    // Italic: *text*
    if (text[i] === '*' && text.slice(i + 1).includes('*') && !text.slice(i).startsWith('**')) {
      if (current) {
        runs.push(new TextRun(current));
        current = '';
      }
      const end = text.indexOf('*', i + 1);
      runs.push(new TextRun({ text: text.slice(i + 1, end), italics: true }));
      i = end + 1;
      continue;
    }

    // Inline code: `text`
    if (text[i] === '`' && text.slice(i + 1).includes('`')) {
      if (current) {
        runs.push(new TextRun(current));
        current = '';
      }
      const end = text.indexOf('`', i + 1);
      runs.push(
        new TextRun({
          text: text.slice(i + 1, end),
          font: 'Courier New',
          color: 'E74C3C',
        })
      );
      i = end + 1;
      continue;
    }

    // Links: [text](url)
    if (text[i] === '[' && text.slice(i + 1).includes('](') && text.slice(i + 1).includes(')')) {
      if (current) {
        runs.push(new TextRun(current));
        current = '';
      }
      const endBracket = text.indexOf(']', i);
      const endParen = text.indexOf(')', endBracket);
      const linkText = text.slice(i + 1, endBracket);
      const url = text.slice(endBracket + 2, endParen);
      runs.push(
        new TextRun({
          text: linkText,
          color: '3498DB',
          underline: { type: UnderlineType.SINGLE },
        })
      );
      i = endParen + 1;
      continue;
    }

    current += text[i];
    i++;
  }

  if (current) {
    runs.push(new TextRun(current));
  }

  return runs.length > 0 ? runs : [new TextRun(text)];
}

/**
 * Write content to a DOCX file
 */
async function writeDocxFile(options: DocxWriteOptions): Promise<DocxWriteResult> {
  const { filePath, content, inputFormat = 'markdown', title, author, overwrite = false } = options;

  try {
    // Check if file exists
    try {
      await fs.access(filePath);
      if (!overwrite) {
        return {
          success: false,
          error: `File already exists: ${filePath}. Set overwrite=true to overwrite.`,
        };
      }
    } catch {
      // File doesn't exist, which is fine
    }

    // Ensure directory exists
    const dir = path.dirname(filePath);
    await fs.mkdir(dir, { recursive: true });

    // Parse content into paragraphs
    const paragraphs = parseMarkdownToParagraphs(content, title);

    // Create document
    const doc = new Document({
      creator: author || 'OpenClaw DOCX Parser',
      title: title || 'Document',
      description: 'Created with OpenClaw DOCX Parser Plugin',
      sections: [
        {
          properties: {},
          children: paragraphs,
        },
      ],
    });

    // Create buffer
    const buffer = await Packer.toBuffer(doc);

    // Write to file
    await fs.writeFile(filePath, buffer);

    // Get file size
    const stats = await fs.stat(filePath);

    return {
      success: true,
      filePath,
      size: stats.size,
    };
  } catch (error: any) {
    return {
      success: false,
      error: error.message || 'Failed to write DOCX file',
    };
  }
}

/**
 * Plugin entry point
 */
export default function register(api: any) {
  // Register the docx_read tool
  api.registerTool({
    name: 'docx_read',
    description:
      '读取并解析 Microsoft Word (.docx) 文件内容。支持提取纯文本、Markdown 或 HTML 格式。',
    parameters: {
      type: 'object',
      properties: {
        filePath: {
          type: 'string',
          description: 'docx 文件的绝对路径',
        },
        outputFormat: {
          type: 'string',
          enum: ['text', 'markdown', 'html'],
          description: '输出格式：text-纯文本, markdown-Markdown格式, html-HTML格式',
          default: 'markdown',
        },
        maxFileSize: {
          type: 'number',
          description: '最大文件大小（MB），默认10MB',
          default: 10,
        },
      },
      required: ['filePath'],
    },
    handler: async (params: any) => {
      return await readDocxFile(params as DocxReadOptions);
    },
  });

  // Register the docx_validate tool
  api.registerTool({
    name: 'docx_validate',
    description: '验证 Microsoft Word (.docx) 文件是否有效。检查文件扩展名和大小。',
    parameters: {
      type: 'object',
      properties: {
        filePath: {
          type: 'string',
          description: 'docx 文件的绝对路径',
        },
        maxFileSize: {
          type: 'number',
          description: '最大文件大小（MB），默认10MB',
          default: 10,
        },
      },
      required: ['filePath'],
    },
    handler: async (params: any) => {
      return await validateDocxFile(params as DocxValidateOptions);
    },
  });

  // Register the docx_write tool
  api.registerTool({
    name: 'docx_write',
    description: '创建并写入 Microsoft Word (.docx) 文件。支持 Markdown 或纯文本格式。',
    parameters: {
      type: 'object',
      properties: {
        filePath: {
          type: 'string',
          description: '输出 docx 文件的绝对路径',
        },
        content: {
          type: 'string',
          description: '要写入的内容（支持 Markdown 或纯文本）',
        },
        inputFormat: {
          type: 'string',
          enum: ['text', 'markdown'],
          description: '输入格式：text-纯文本, markdown-Markdown格式（默认）',
          default: 'markdown',
        },
        title: {
          type: 'string',
          description: '文档标题（可选）',
        },
        author: {
          type: 'string',
          description: '作者名称（可选）',
        },
        overwrite: {
          type: 'boolean',
          description: '是否覆盖已存在的文件（默认：false）',
          default: false,
        },
      },
      required: ['filePath', 'content'],
    },
    handler: async (params: any) => {
      return await writeDocxFile(params as DocxWriteOptions);
    },
  });

  // Register CLI commands
  api.registerCli(
    ({ program }: any) => {
      program
        .command('docx-read <file>')
        .description('Read and display DOCX file content')
        .option('-f, --format <format>', 'Output format (text|markdown|html)', 'markdown')
        .option('-s, --max-size <size>', 'Maximum file size in MB', '10')
        .action(async (file: string, options: any) => {
          const result = await readDocxFile({
            filePath: path.resolve(file),
            outputFormat: options.format,
            maxFileSize: parseInt(options.maxSize, 10),
          });

          if (result.success) {
            console.log(result.content);
            if (result.warnings && result.warnings.length > 0) {
              console.warn('\nWarnings:', result.warnings);
            }
          } else {
            console.error('Error:', result.error);
            process.exit(1);
          }
        });

      program
        .command('docx-validate <file>')
        .description('Validate a DOCX file')
        .option('-s, --max-size <size>', 'Maximum file size in MB', '10')
        .action(async (file: string, options: any) => {
          const result = await validateDocxFile({
            filePath: path.resolve(file),
            maxFileSize: parseInt(options.maxSize, 10),
          });

          if (result.success) {
            if (result.valid) {
              console.log('✓ File is valid');
              if (result.size !== undefined) {
                console.log(`  Size: ${(result.size / 1024 / 1024).toFixed(2)} MB`);
              }
            } else {
              console.error('✗ File is invalid:', result.error);
              process.exit(1);
            }
          } else {
            console.error('Error:', result.error);
            process.exit(1);
          }
        });

      program
        .command('docx-write <file> [content]')
        .description('Create a DOCX file from markdown or text')
        .option('-f, --format <format>', 'Input format (text|markdown)', 'markdown')
        .option('-t, --title <title>', 'Document title')
        .option('-a, --author <author>', 'Author name')
        .option('-o, --overwrite', 'Overwrite existing file')
        .option('-c, --content <content>', 'Content (alternative to positional arg)')
        .action(async (file: string, contentArg: string, options: any) => {
          const content = options.content || contentArg;
          if (!content) {
            console.error('Error: Content is required. Provide it as argument or use --content');
            process.exit(1);
          }

          const result = await writeDocxFile({
            filePath: path.resolve(file),
            content,
            inputFormat: options.format,
            title: options.title,
            author: options.author,
            overwrite: options.overwrite || false,
          });

          if (result.success) {
            console.log(`✓ DOCX file created: ${result.filePath}`);
            if (result.size !== undefined) {
              console.log(`  Size: ${(result.size / 1024 / 1024).toFixed(2)} MB`);
            }
          } else {
            console.error('Error:', result.error);
            process.exit(1);
          }
        });
    },
    { commands: ['docx-read', 'docx-validate', 'docx-write'] }
  );

  api.logger.info('DOCX Parser plugin loaded');
}

// Export for testing
export { readDocxFile, validateDocxFile, writeDocxFile, htmlToMarkdown };
