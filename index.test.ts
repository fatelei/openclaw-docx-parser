/**
 * Unit tests for DOCX Parser plugin
 */

import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { readDocxFile, validateDocxFile, htmlToMarkdown } from './index';
import * as fs from 'fs/promises';
import * as mammoth from 'mammoth';

// Mock fs module
vi.mock('fs/promises');

// Mock mammoth module
vi.mock('mammoth', () => ({
  convertToHtml: vi.fn(),
  extractRawText: vi.fn(),
}));

describe('htmlToMarkdown', () => {
  it('should convert h1-h6 headings to markdown', () => {
    expect(htmlToMarkdown('<h1>Title</h1>')).toBe('# Title');
    expect(htmlToMarkdown('<h2>Subtitle</h2>')).toBe('## Subtitle');
    expect(htmlToMarkdown('<h3>Heading 3</h3>')).toBe('### Heading 3');
  });

  it('should convert bold and italic to markdown', () => {
    expect(htmlToMarkdown('<strong>Bold</strong>')).toBe('**Bold**');
    expect(htmlToMarkdown('<b>Bold</b>')).toBe('**Bold**');
    expect(htmlToMarkdown('<em>Italic</em>')).toBe('*Italic*');
    expect(htmlToMarkdown('<i>Italic</i>')).toBe('*Italic*');
  });

  it('should convert links to markdown', () => {
    expect(htmlToMarkdown('<a href="https://example.com">Link</a>')).toBe('[Link](https://example.com)');
  });

  it('should convert images to markdown', () => {
    expect(htmlToMarkdown('<img src="image.jpg" alt="Alt text">')).toBe('![Alt text](image.jpg)');
  });

  it('should convert line breaks', () => {
    expect(htmlToMarkdown('Line 1<br>Line 2')).toBe('Line 1\nLine 2');
  });

  it('should convert paragraphs', () => {
    expect(htmlToMarkdown('<p>Paragraph 1</p><p>Paragraph 2</p>')).toBe('Paragraph 1\n\nParagraph 2');
  });

  it('should convert lists', () => {
    expect(htmlToMarkdown('<ul><li>Item 1</li><li>Item 2</li></ul>')).toContain('- Item 1');
    expect(htmlToMarkdown('<ul><li>Item 1</li><li>Item 2</li></ul>')).toContain('- Item 2');
  });

  it('should convert code blocks', () => {
    expect(htmlToMarkdown('<pre><code>const x = 1;</code></pre>')).toBe('```\nconst x = 1;\n```');
    expect(htmlToMarkdown('<code>inline</code>')).toBe('`inline`');
  });

  it('should convert blockquotes', () => {
    expect(htmlToMarkdown('<blockquote>Quote</blockquote>')).toBe('> Quote');
  });

  it('should convert horizontal rules', () => {
    expect(htmlToMarkdown('<hr>')).toBe('---');
  });

  it('should remove remaining HTML tags', () => {
    expect(htmlToMarkdown('<span>Text</span>')).toBe('Text');
  });

  it('should clean up extra whitespace', () => {
    expect(htmlToMarkdown('<p>A</p>\n\n\n\n<p>B</p>')).toBe('A\n\nB');
  });

  it('should handle complex HTML', () => {
    const html = '<h1>Document Title</h1><p>This is <strong>bold</strong> and <em>italic</em> text.</p><ul><li>Item 1</li><li>Item 2</li></ul>';
    const result = htmlToMarkdown(html);
    expect(result).toContain('# Document Title');
    expect(result).toContain('**bold**');
    expect(result).toContain('*italic*');
    expect(result).toContain('- Item 1');
  });
});

describe('validateDocxFile', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('should validate a valid .docx file', async () => {
    vi.mocked(fs.stat).mockResolvedValue({ isFile: () => true, size: 1024 * 1024 } as any);

    const result = await validateDocxFile({
      filePath: '/test/document.docx',
      maxFileSize: 10,
    });

    expect(result.success).toBe(true);
    expect(result.valid).toBe(true);
    expect(result.size).toBe(1024 * 1024);
  });

  it('should reject non-.docx extension', async () => {
    vi.mocked(fs.stat).mockResolvedValue({ isFile: () => true } as any);

    const result = await validateDocxFile({
      filePath: '/test/document.txt',
    });

    expect(result.success).toBe(true);
    expect(result.valid).toBe(false);
    expect(result.error).toContain('.docx extension');
  });

  it('should reject files that are too large', async () => {
    vi.mocked(fs.stat).mockResolvedValue({ isFile: () => true, size: 15 * 1024 * 1024 } as any);

    const result = await validateDocxFile({
      filePath: '/test/document.docx',
      maxFileSize: 10,
    });

    expect(result.success).toBe(true);
    expect(result.valid).toBe(false);
    expect(result.error).toContain('exceeds maximum');
    expect(result.error).toContain('10MB');
  });

  it('should reject paths that are not files', async () => {
    vi.mocked(fs.stat).mockResolvedValue({ isFile: () => false } as any);

    const result = await validateDocxFile({
      filePath: '/test/document.docx',
    });

    expect(result.success).toBe(true);
    expect(result.valid).toBe(false);
    expect(result.error).toContain('not a file');
  });

  it('should handle file system errors', async () => {
    vi.mocked(fs.stat).mockRejectedValue(new Error('File not found'));

    const result = await validateDocxFile({
      filePath: '/test/document.docx',
    });

    expect(result.success).toBe(false);
    expect(result.valid).toBe(false);
    expect(result.error).toBe('File not found');
  });

  it('should handle case-insensitive file extension', async () => {
    vi.mocked(fs.stat).mockResolvedValue({ isFile: () => true, size: 1024 } as any);

    const resultLower = await validateDocxFile({ filePath: '/test/document.docx' });
    const resultUpper = await validateDocxFile({ filePath: '/test/document.DOCX' });

    expect(resultLower.valid).toBe(true);
    expect(resultUpper.valid).toBe(true);
  });
});

describe('readDocxFile', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('should read a .docx file as markdown', async () => {
    vi.mocked(fs.stat).mockResolvedValue({ isFile: () => true, size: 1024 } as any);
    vi.mocked(mammoth.convertToHtml).mockResolvedValue({
      value: '<h1>Title</h1><p>Content</p>',
      messages: [],
    });

    const result = await readDocxFile({
      filePath: '/test/document.docx',
      outputFormat: 'markdown',
    });

    expect(result.success).toBe(true);
    expect(result.content).toContain('# Title');
    expect(result.content).toContain('Content');
  });

  it('should read a .docx file as HTML', async () => {
    vi.mocked(fs.stat).mockResolvedValue({ isFile: () => true, size: 1024 } as any);
    vi.mocked(mammoth.convertToHtml).mockResolvedValue({
      value: '<h1>Title</h1><p>Content</p>',
      messages: [],
    });

    const result = await readDocxFile({
      filePath: '/test/document.docx',
      outputFormat: 'html',
    });

    expect(result.success).toBe(true);
    expect(result.content).toBe('<h1>Title</h1><p>Content</p>');
  });

  it('should read a .docx file as plain text', async () => {
    vi.mocked(fs.stat).mockResolvedValue({ isFile: () => true, size: 1024 } as any);
    vi.mocked(mammoth.extractRawText).mockResolvedValue({
      value: 'Plain text content',
      messages: [],
    });

    const result = await readDocxFile({
      filePath: '/test/document.docx',
      outputFormat: 'text',
    });

    expect(result.success).toBe(true);
    expect(result.content).toBe('Plain text content');
  });

  it('should default to markdown format', async () => {
    vi.mocked(fs.stat).mockResolvedValue({ isFile: () => true, size: 1024 } as any);
    vi.mocked(mammoth.convertToHtml).mockResolvedValue({
      value: '<p>Content</p>',
      messages: [],
    });

    const result = await readDocxFile({
      filePath: '/test/document.docx',
    });

    expect(result.success).toBe(true);
    expect(mammoth.convertToHtml).toHaveBeenCalled();
  });

  it('should include warnings from mammoth', async () => {
    vi.mocked(fs.stat).mockResolvedValue({ isFile: () => true, size: 1024 } as any);
    vi.mocked(mammoth.convertToHtml).mockResolvedValue({
      value: '<p>Content</p>',
      messages: ['Warning: Unsupported element'],
    });

    const result = await readDocxFile({
      filePath: '/test/document.docx',
      outputFormat: 'html',
    });

    expect(result.success).toBe(true);
    expect(result.warnings).toEqual(['Warning: Unsupported element']);
  });

  it('should handle invalid files', async () => {
    vi.mocked(fs.stat).mockResolvedValue({ isFile: () => true, size: 1024 } as any);

    const result = await readDocxFile({
      filePath: '/test/document.txt',
    });

    expect(result.success).toBe(false);
    expect(result.error).toContain('.docx extension');
  });

  it('should handle mammoth errors', async () => {
    vi.mocked(fs.stat).mockResolvedValue({ isFile: () => true, size: 1024 } as any);
    vi.mocked(mammoth.convertToHtml).mockRejectedValue(new Error('Corrupt DOCX'));

    const result = await readDocxFile({
      filePath: '/test/document.docx',
      outputFormat: 'markdown',
    });

    expect(result.success).toBe(false);
    expect(result.error).toBe('Corrupt DOCX');
  });

  it('should handle large files according to maxFileSize', async () => {
    vi.mocked(fs.stat).mockResolvedValue({ isFile: () => true, size: 20 * 1024 * 1024 } as any);

    const result = await readDocxFile({
      filePath: '/test/document.docx',
      maxFileSize: 10,
    });

    expect(result.success).toBe(false);
    expect(result.error).toContain('exceeds maximum');
  });

  it('should allow files within size limit', async () => {
    vi.mocked(fs.stat).mockResolvedValue({ isFile: () => true, size: 5 * 1024 * 1024 } as any);
    vi.mocked(mammoth.convertToHtml).mockResolvedValue({
      value: '<p>Content</p>',
      messages: [],
    });

    const result = await readDocxFile({
      filePath: '/test/document.docx',
      maxFileSize: 10,
    });

    expect(result.success).toBe(true);
  });

  it('should return empty warnings array if no warnings', async () => {
    vi.mocked(fs.stat).mockResolvedValue({ isFile: () => true, size: 1024 } as any);
    vi.mocked(mammoth.convertToHtml).mockResolvedValue({
      value: '<p>Content</p>',
      messages: [],
    });

    const result = await readDocxFile({
      filePath: '/test/document.docx',
    });

    expect(result.success).toBe(true);
    expect(result.warnings).toBeUndefined();
  });
});
