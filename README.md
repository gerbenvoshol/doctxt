# doctxt

doctxt is a simple, fast bidirectional conversion tool written in C:
- **doctxt**: Convert docx to txt
- **md2docx**: Convert markdown to docx

## Dependencies

### For building

- libxml2 (for doctxt)
- libzip (for doctxt)
- libmd4c (for md2docx)

```sh
$ apt install libxml2-dev libzip-dev libmd4c-dev
```

### Installation

Install dependencies first.

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
