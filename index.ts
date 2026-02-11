/**
 * OpenClaw DOCX Parser Plugin
 *
 * A plugin for reading and parsing Microsoft Word (.docx) files.
 * Supports extracting text, markdown, and HTML content from DOCX files.
 */

import * as mammoth from 'mammoth';
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
    },
    { commands: ['docx-read', 'docx-validate'] }
  );

  api.logger.info('DOCX Parser plugin loaded');
}

// Export for testing
export { readDocxFile, validateDocxFile, htmlToMarkdown };
