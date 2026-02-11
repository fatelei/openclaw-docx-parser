/**
 * Unit tests for DOCX Parser plugin
 */

import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { readDocxFile, validateDocxFile, writeDocxFile, htmlToMarkdown } from './index';
import * as fs from 'fs/promises';
import * as mammoth from 'mammoth';

// Mock fs module
vi.mock('fs/promises');

// Mock mammoth module
vi.mock('mammoth', () => ({
  convertToHtml: vi.fn(),
  extractRawText: vi.fn(),
}));

// Mock docx module
vi.mock('docx', () => ({
  Document: vi.fn().mockImplementation(() => ({})),
  Packer: {
    toBuffer: vi.fn().mockResolvedValue(Buffer.from('mock docx content')),
  },
  Paragraph: vi.fn().mockImplementation(() => ({})),
  TextRun: vi.fn().mockImplementation(() => ({})),
  HeadingLevel: {
    TITLE: 'TITLE',
    HEADING_1: 'HEADING_1',
    HEADING_2: 'HEADING_2',
    HEADING_3: 'HEADING_3',
    HEADING_4: 'HEADING_4',
    HEADING_5: 'HEADING_5',
    HEADING_6: 'HEADING_6',
  },
  AlignmentType: {
    CENTER: 'CENTER',
  },
  UnderlineType: {
    SINGLE: 'SINGLE',
  },
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

describe('writeDocxFile', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('should create a new DOCX file', async () => {
    vi.mocked(fs.access).mockRejectedValue(new Error('File not found'));
    vi.mocked(fs.mkdir).mockResolvedValue(undefined);
    vi.mocked(fs.writeFile).mockResolvedValue(undefined);
    vi.mocked(fs.stat).mockResolvedValue({ size: 2048 } as any);

    const result = await writeDocxFile({
      filePath: '/test/output.docx',
      content: '# Test Document\n\nThis is a test.',
    });

    expect(result.success).toBe(true);
    expect(result.filePath).toBe('/test/output.docx');
    expect(result.size).toBe(2048);
  });

  it('should fail if file exists and overwrite is false', async () => {
    vi.mocked(fs.access).mockResolvedValue(undefined);

    const result = await writeDocxFile({
      filePath: '/test/existing.docx',
      content: 'Content',
      overwrite: false,
    });

    expect(result.success).toBe(false);
    expect(result.error).toContain('already exists');
  });

  it('should overwrite existing file when overwrite is true', async () => {
    vi.mocked(fs.access).mockResolvedValue(undefined);
    vi.mocked(fs.mkdir).mockResolvedValue(undefined);
    vi.mocked(fs.writeFile).mockResolvedValue(undefined);
    vi.mocked(fs.stat).mockResolvedValue({ size: 2048 } as any);

    const result = await writeDocxFile({
      filePath: '/test/existing.docx',
      content: 'New content',
      overwrite: true,
    });

    expect(result.success).toBe(true);
    expect(fs.writeFile).toHaveBeenCalled();
  });

  it('should create directory if it does not exist', async () => {
    vi.mocked(fs.access).mockRejectedValue(new Error('File not found'));
    vi.mocked(fs.mkdir).mockResolvedValue(undefined);
    vi.mocked(fs.writeFile).mockResolvedValue(undefined);
    vi.mocked(fs.stat).mockResolvedValue({ size: 2048 } as any);

    const result = await writeDocxFile({
      filePath: '/test/newdir/output.docx',
      content: 'Content',
    });

    expect(result.success).toBe(true);
    expect(fs.mkdir).toHaveBeenCalledWith('/test/newdir', { recursive: true });
  });

  it('should include title and author in document', async () => {
    vi.mocked(fs.access).mockRejectedValue(new Error('File not found'));
    vi.mocked(fs.mkdir).mockResolvedValue(undefined);
    vi.mocked(fs.writeFile).mockResolvedValue(undefined);
    vi.mocked(fs.stat).mockResolvedValue({ size: 2048 } as any);

    const result = await writeDocxFile({
      filePath: '/test/output.docx',
      content: 'Content',
      title: 'My Document',
      author: 'Test Author',
    });

    expect(result.success).toBe(true);
  });

  it('should handle write errors', async () => {
    vi.mocked(fs.access).mockRejectedValue(new Error('File not found'));
    vi.mocked(fs.mkdir).mockResolvedValue(undefined);
    vi.mocked(fs.writeFile).mockRejectedValue(new Error('Write failed'));

    const result = await writeDocxFile({
      filePath: '/test/output.docx',
      content: 'Content',
    });

    expect(result.success).toBe(false);
    expect(result.error).toBe('Write failed');
  });

  it('should handle markdown content with headings', async () => {
    vi.mocked(fs.access).mockRejectedValue(new Error('File not found'));
    vi.mocked(fs.mkdir).mockResolvedValue(undefined);
    vi.mocked(fs.writeFile).mockResolvedValue(undefined);
    vi.mocked(fs.stat).mockResolvedValue({ size: 2048 } as any);

    const content = '# Heading 1\n\n## Heading 2\n\n### Heading 3';
    const result = await writeDocxFile({
      filePath: '/test/output.docx',
      content,
    });

    expect(result.success).toBe(true);
  });

  it('should handle markdown content with lists', async () => {
    vi.mocked(fs.access).mockRejectedValue(new Error('File not found'));
    vi.mocked(fs.mkdir).mockResolvedValue(undefined);
    vi.mocked(fs.writeFile).mockResolvedValue(undefined);
    vi.mocked(fs.stat).mockResolvedValue({ size: 2048 } as any);

    const content = '- Item 1\n- Item 2\n- Item 3';
    const result = await writeDocxFile({
      filePath: '/test/output.docx',
      content,
    });

    expect(result.success).toBe(true);
  });

  it('should handle markdown content with bold and italic', async () => {
    vi.mocked(fs.access).mockRejectedValue(new Error('File not found'));
    vi.mocked(fs.mkdir).mockResolvedValue(undefined);
    vi.mocked(fs.writeFile).mockResolvedValue(undefined);
    vi.mocked(fs.stat).mockResolvedValue({ size: 2048 } as any);

    const content = '**Bold text** and *italic text*';
    const result = await writeDocxFile({
      filePath: '/test/output.docx',
      content,
    });

    expect(result.success).toBe(true);
  });

  it('should handle markdown content with code', async () => {
    vi.mocked(fs.access).mockRejectedValue(new Error('File not found'));
    vi.mocked(fs.mkdir).mockResolvedValue(undefined);
    vi.mocked(fs.writeFile).mockResolvedValue(undefined);
    vi.mocked(fs.stat).mockResolvedValue({ size: 2048 } as any);

    const content = 'Inline `code` and regular text';
    const result = await writeDocxFile({
      filePath: '/test/output.docx',
      content,
    });

    expect(result.success).toBe(true);
  });

  it('should handle markdown content with blockquotes', async () => {
    vi.mocked(fs.access).mockRejectedValue(new Error('File not found'));
    vi.mocked(fs.mkdir).mockResolvedValue(undefined);
    vi.mocked(fs.writeFile).mockResolvedValue(undefined);
    vi.mocked(fs.stat).mockResolvedValue({ size: 2048 } as any);

    const content = '> This is a quote';
    const result = await writeDocxFile({
      filePath: '/test/output.docx',
      content,
    });

    expect(result.success).toBe(true);
  });

  it('should handle markdown content with horizontal rules', async () => {
    vi.mocked(fs.access).mockRejectedValue(new Error('File not found'));
    vi.mocked(fs.mkdir).mockResolvedValue(undefined);
    vi.mocked(fs.writeFile).mockResolvedValue(undefined);
    vi.mocked(fs.stat).mockResolvedValue({ size: 2048 } as any);

    const content = 'Text before\n\n---\n\nText after';
    const result = await writeDocxFile({
      filePath: '/test/output.docx',
      content,
    });

    expect(result.success).toBe(true);
  });
});
