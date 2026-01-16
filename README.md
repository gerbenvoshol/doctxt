# doctxt

doctxt is a simple, fast bidirectional conversion tool written in C:
- **doctxt**: Convert docx to txt
- **md2docx**: Convert markdown to docx
- **docx2md**: Convert docx to markdown

## Dependencies

### For building

No external libraries required! This project uses embedded dependencies:
- **txml** (embedded) - for XML parsing in doctxt and docx2md
- **miniz** (embedded) - for ZIP handling in all tools
- **md4c** (embedded) - for Markdown parsing in md2docx

Just a C compiler is needed.

### Installation

```sh
$ make clean
$ make
$ make install
```

## Usage

### doctxt - DOCX to Text Converter

```sh
$ doctxt [FILE] [-o OUTFILE] [-c]
```

**Options:**
- `-o OUTFILE`: Specify the output file (default: out.txt)
- `-c`: Extract only comments from the document
- `-v`: Display version information

If -o is omitted, output will be written to out.txt

**Features:**
- Extract text content from docx files
- Extract tables (preserves table structure with tab-separated columns)
- Extract comments with author attribution
- Fast and lightweight C implementation

### md2docx - Markdown to DOCX Converter

Convert Markdown files to Microsoft Word DOCX format.

```sh
$ md2docx input.md [-o output.docx]
```

**Options:**
- `-o FILE`: Specify output file (default: output.docx)
- `-v`: Display version information
- `-h`: Display help message

**Supported Markdown Features:**
- **Headings** (# through ######)
- **Text formatting**: **bold**, *italic*, `code`, ~~strikethrough~~, <u>underline</u>
- **Lists**: Unordered (bullet) and ordered (numbered) lists
- **Code blocks**: Fenced code blocks with syntax highlighting info
- **Tables**: Full table support with alignment
- **Links**: Hyperlinks (rendered as underlined text)
- **Images**: Embedded images (reads from filesystem)
- **Horizontal rules** (---)
- **GitHub-flavored Markdown**: Tables, strikethrough, task lists

**Example:**

```sh
# Convert README.md to Word document
$ md2docx README.md -o documentation.docx

# Quick conversion with default output name
$ md2docx notes.md
# Creates output.docx
```

**Image Support:**

Images referenced in markdown will be embedded into the DOCX file:

```markdown
![Alt text](path/to/image.png)
```

The tool will read the image file and embed it directly into the Word document.

### docx2md - DOCX to Markdown Converter

Convert Microsoft Word DOCX files to Markdown format.

```sh
$ docx2md input.docx [-o output.md]
```

**Options:**
- `-o FILE`: Specify output file (default: output.md)
- `-v`: Display version information
- `-h`: Display help message

**Supported DOCX Features:**
- **Headings** (Heading1 through Heading6)
- **Text formatting**: **bold**, *italic*, `code`, ~~strikethrough~~
- **Code blocks**: Paragraphs with "Code" style
- **Tables**: Full table support with headers
- **Line breaks**: Manual line breaks within paragraphs
- **Horizontal rules**: Paragraph borders

**Example:**

```sh
# Convert Word document to Markdown
$ docx2md document.docx -o README.md

# Quick conversion with default output name
$ docx2md notes.docx
# Creates output.md
```

**Note:** 
- The tool extracts text content and formatting from DOCX files
- Hyperlinks are converted to plain text (link text without URLs, as URLs may not be stored in simple DOCX files)
- Images are not yet extracted in this version
