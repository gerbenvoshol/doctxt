# doctxt

doctxt is a simple, fast bidirectional conversion tool written in C:
- **doctxt**: Convert docx to txt
- **md2docx**: Convert markdown to docx
- **docx2md**: Convert docx to markdown
- **xlsx2csv**: Convert xlsx (Excel 2007+) to CSV
- **xlsx2md**: Convert xlsx (Excel 2007+) to Markdown tables

## Dependencies

### For building

No external libraries required! This project uses embedded dependencies:
- **txml** (embedded) - for XML parsing in doctxt, docx2md, xlsx2csv, and xlsx2md
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
- **Images**: Embedded images are extracted to the output directory

**Example:**

```sh
# Convert Word document to Markdown
$ docx2md document.docx -o README.md

# Quick conversion with default output name
$ docx2md notes.docx
# Creates output.md
```

**Image Support:**

Images embedded in the DOCX file are automatically extracted to the same directory as the output markdown file and referenced using markdown image syntax:

```markdown
![Alt text](image1.png)
```

**Note:** 
- The tool extracts text content and formatting from DOCX files
- Hyperlinks are converted to plain text (link text without URLs, as URLs may not be stored in simple DOCX files)
- Images are extracted from the DOCX archive and saved to the output directory

### xlsx2csv - XLSX to CSV Converter

Convert Microsoft Excel 2007+ XLSX files to CSV format.

```sh
$ xlsx2csv -if input.xlsx [-sh sheet_number] [-of output.csv]
```

**Options:**
- `-if FILE`: Input spreadsheet in Excel 2007 format (required)
- `-sh NUMBER`: Sheet number to convert (default: 1)
- `-of FILE`: Output CSV file (default: stdout)
- `-v`: Display version information
- `-h`: Display help message

**Example:**

```sh
# Convert first sheet to CSV
$ xlsx2csv -if data.xlsx -of data.csv

# Convert specific sheet to CSV
$ xlsx2csv -if data.xlsx -sh 2 -of sheet2.csv

# Output to stdout (can be piped)
$ xlsx2csv -if data.xlsx | grep "pattern"
```

**Features:**
- Fast conversion using embedded txml and miniz libraries
- Handles shared strings for efficient storage
- Properly escapes CSV fields containing commas, quotes, or newlines
- Supports all Excel 2007+ XLSX files

### xlsx2md - XLSX to Markdown Converter

Convert Microsoft Excel 2007+ XLSX files to Markdown tables.

```sh
$ xlsx2md -if input.xlsx [-sh sheet_number] [-of output.md]
```

**Options:**
- `-if FILE`: Input spreadsheet in Excel 2007 format (required)
- `-sh NUMBER`: Sheet number to convert (default: 1)
- `-of FILE`: Output Markdown file (default: stdout)
- `-v`: Display version information
- `-h`: Display help message

**Example:**

```sh
# Convert first sheet to Markdown table
$ xlsx2md -if data.xlsx -of data.md

# Convert specific sheet to Markdown
$ xlsx2md -if data.xlsx -sh 2 -of sheet2.md

# Output to stdout
$ xlsx2md -if data.xlsx
```

**Features:**
- Creates properly formatted Markdown tables
- First row is automatically treated as table header
- Escapes pipe characters in cell content
- Handles empty cells gracefully
- Perfect for including Excel data in documentation

