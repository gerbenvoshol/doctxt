/*
 * docx2md - Convert DOCX to Markdown
 * 
 * Uses txml for XML parsing and miniz for ZIP handling
 */

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>

#include "miniz.h"
#include "util.h"

#define TXML_DEFINE
#include "txml.h"

#define VERSION_STR "0.1"
#define TEMPFILE "/tmp/docx2md-temp.xml"
#define MAX_BUFFER_SIZE (10 * 1024 * 1024)

/* Context structure to hold state during conversion */
typedef struct {
    FILE *output;
    int in_bold;
    int in_italic;
    int in_code;
    int in_strikethrough;
    int in_underline;
    int in_table;
    int table_col_count;
    int first_table_row;
} md_context;

/* Forward declarations */
static void process_paragraph(struct txml_node *para, md_context *ctx);
static void process_run(struct txml_node *run, md_context *ctx);
static void process_table(struct txml_node *table, md_context *ctx);
static char *xml_unescape(const char *in);

/* XML unescape helper */
static char *xml_unescape(const char *in) {
    if (!in) return NULL;
    
    char *out = malloc(strlen(in) * 2);
    if (!out) return NULL;
    
    char *dst = out;
    const char *src = in;

    while (*src) {
        if (!strncmp(src, "&gt;", 4)) {
            *dst++ = '>';
            src += 4;
        } else if (!strncmp(src, "&lt;", 4)) {
            *dst++ = '<';
            src += 4;
        } else if (!strncmp(src, "&amp;", 5)) {
            *dst++ = '&';
            src += 5;
        } else if (!strncmp(src, "&quot;", 6)) {
            *dst++ = '"';
            src += 6;
        } else if (!strncmp(src, "&apos;", 6)) {
            *dst++ = '\'';
            src += 6;
        } else {
            *dst++ = *src++;
        }
    }
    *dst = '\0';
    return out;
}

/* Get paragraph style to determine heading level or code block */
static const char *get_paragraph_style(struct txml_node *para) {
    struct txml_node *pPr = txml_find(para, NULL, TXML_ELEMENT, "w:pPr", NULL, 0);
    if (!pPr) return NULL;
    
    struct txml_node *pStyle = txml_find(pPr, NULL, TXML_ELEMENT, "w:pStyle", NULL, 0);
    if (!pStyle) return NULL;
    
    struct txml_node *val_attr = txml_find(pStyle, NULL, TXML_ATTRIBUTE, "w:val", NULL, 0);
    if (!val_attr) return NULL;
    
    return val_attr->value;
}

/* Check if paragraph has horizontal rule */
static int has_horizontal_rule(struct txml_node *para) {
    struct txml_node *pPr = txml_find(para, NULL, TXML_ELEMENT, "w:pPr", NULL, 0);
    if (!pPr) return 0;
    
    struct txml_node *pBdr = txml_find(pPr, NULL, TXML_ELEMENT, "w:pBdr", NULL, 0);
    if (!pBdr) return 0;
    
    struct txml_node *bottom = txml_find(pBdr, NULL, TXML_ELEMENT, "w:bottom", NULL, 0);
    return bottom != NULL;
}

/* Check if text run has formatting */
static void check_run_formatting(struct txml_node *run, md_context *ctx) {
    struct txml_node *rPr = txml_find(run, NULL, TXML_ELEMENT, "w:rPr", NULL, 0);
    if (!rPr) return;
    
    ctx->in_bold = txml_find(rPr, NULL, TXML_ELEMENT, "w:b", NULL, 0) != NULL;
    ctx->in_italic = txml_find(rPr, NULL, TXML_ELEMENT, "w:i", NULL, 0) != NULL;
    ctx->in_strikethrough = txml_find(rPr, NULL, TXML_ELEMENT, "w:strike", NULL, 0) != NULL;
    ctx->in_underline = txml_find(rPr, NULL, TXML_ELEMENT, "w:u", NULL, 0) != NULL;
    
    struct txml_node *rStyle = txml_find(rPr, NULL, TXML_ELEMENT, "w:rStyle", NULL, 0);
    if (rStyle) {
        struct txml_node *val_attr = txml_find(rStyle, NULL, TXML_ATTRIBUTE, "w:val", NULL, 0);
        if (val_attr && val_attr->value && strcmp(val_attr->value, "CodeChar") == 0) {
            ctx->in_code = 1;
        }
    }
}

/* Process a text run (w:r) */
static void process_run(struct txml_node *run, md_context *ctx) {
    int old_bold = ctx->in_bold;
    int old_italic = ctx->in_italic;
    int old_code = ctx->in_code;
    int old_strike = ctx->in_strikethrough;
    
    check_run_formatting(run, ctx);
    
    /* Open formatting markers */
    if (ctx->in_strikethrough && !old_strike) fprintf(ctx->output, "~~");
    if (ctx->in_bold && !old_bold) fprintf(ctx->output, "**");
    if (ctx->in_italic && !old_italic) fprintf(ctx->output, "*");
    if (ctx->in_code && !old_code) fprintf(ctx->output, "`");
    
    /* Extract text content */
    struct txml_node *text_node = txml_find(run, NULL, TXML_ELEMENT, "w:t", NULL, 0);
    if (text_node && text_node->value) {
        char *unescaped = xml_unescape(text_node->value);
        if (unescaped) {
            fprintf(ctx->output, "%s", unescaped);
            free(unescaped);
        }
    }
    
    /* Check for line break */
    struct txml_node *br = txml_find(run, NULL, TXML_ELEMENT, "w:br", NULL, 0);
    if (br) {
        fprintf(ctx->output, "  \n");
    }
    
    /* Close formatting markers in reverse order */
    if (ctx->in_code && !old_code) fprintf(ctx->output, "`");
    if (ctx->in_italic && !old_italic) fprintf(ctx->output, "*");
    if (ctx->in_bold && !old_bold) fprintf(ctx->output, "**");
    if (ctx->in_strikethrough && !old_strike) fprintf(ctx->output, "~~");
    
    /* Reset to old state */
    ctx->in_bold = old_bold;
    ctx->in_italic = old_italic;
    ctx->in_code = old_code;
    ctx->in_strikethrough = old_strike;
}

/* Process a paragraph (w:p) */
static void process_paragraph(struct txml_node *para, md_context *ctx) {
    const char *style = get_paragraph_style(para);
    
    /* Handle headings */
    if (style) {
        if (strcmp(style, "Heading1") == 0) {
            fprintf(ctx->output, "# ");
        } else if (strcmp(style, "Heading2") == 0) {
            fprintf(ctx->output, "## ");
        } else if (strcmp(style, "Heading3") == 0) {
            fprintf(ctx->output, "### ");
        } else if (strcmp(style, "Heading4") == 0) {
            fprintf(ctx->output, "#### ");
        } else if (strcmp(style, "Heading5") == 0) {
            fprintf(ctx->output, "##### ");
        } else if (strcmp(style, "Heading6") == 0) {
            fprintf(ctx->output, "###### ");
        } else if (strcmp(style, "Code") == 0) {
            /* Code block - process differently */
            fprintf(ctx->output, "```\n");
            struct txml_node *run = NULL;
            while ((run = txml_find(para, run, TXML_ELEMENT, "w:r", NULL, 0))) {
                struct txml_node *text_node = txml_find(run, NULL, TXML_ELEMENT, "w:t", NULL, 0);
                if (text_node && text_node->value) {
                    char *unescaped = xml_unescape(text_node->value);
                    if (unescaped) {
                        fprintf(ctx->output, "%s", unescaped);
                        free(unescaped);
                    }
                }
            }
            fprintf(ctx->output, "\n```\n\n");
            return;
        }
    }
    
    /* Check for horizontal rule */
    if (has_horizontal_rule(para)) {
        fprintf(ctx->output, "---\n\n");
        return;
    }
    
    /* Process all runs in the paragraph */
    struct txml_node *run = NULL;
    int has_content = 0;
    while ((run = txml_find(para, run, TXML_ELEMENT, "w:r", NULL, 0))) {
        process_run(run, ctx);
        has_content = 1;
    }
    
    /* End paragraph with double newline if it had content */
    if (has_content || style) {
        fprintf(ctx->output, "\n\n");
    }
}

/* Process a table (w:tbl) */
static void process_table(struct txml_node *table, md_context *ctx) {
    ctx->in_table = 1;
    ctx->first_table_row = 1;
    ctx->table_col_count = 0;
    
    /* First pass: count columns */
    struct txml_node *first_row = txml_find(table, NULL, TXML_ELEMENT, "w:tr", NULL, 0);
    if (first_row) {
        struct txml_node *cell = NULL;
        while ((cell = txml_find(first_row, cell, TXML_ELEMENT, "w:tc", NULL, 0))) {
            ctx->table_col_count++;
        }
    }
    
    /* Process all rows */
    struct txml_node *row = NULL;
    while ((row = txml_find(table, row, TXML_ELEMENT, "w:tr", NULL, 0))) {
        fprintf(ctx->output, "|");
        
        struct txml_node *cell = NULL;
        while ((cell = txml_find(row, cell, TXML_ELEMENT, "w:tc", NULL, 0))) {
            fprintf(ctx->output, " ");
            
            /* Process all paragraphs in the cell */
            struct txml_node *para = NULL;
            int first_para = 1;
            while ((para = txml_find(cell, para, TXML_ELEMENT, "w:p", NULL, 0))) {
                if (!first_para) fprintf(ctx->output, " ");
                first_para = 0;
                
                struct txml_node *run = NULL;
                while ((run = txml_find(para, run, TXML_ELEMENT, "w:r", NULL, 0))) {
                    process_run(run, ctx);
                }
            }
            
            fprintf(ctx->output, " |");
        }
        
        fprintf(ctx->output, "\n");
        
        /* Add separator row after header */
        if (ctx->first_table_row) {
            fprintf(ctx->output, "|");
            for (int i = 0; i < ctx->table_col_count; i++) {
                fprintf(ctx->output, "---------|");
            }
            fprintf(ctx->output, "\n");
            ctx->first_table_row = 0;
        }
    }
    
    fprintf(ctx->output, "\n");
    ctx->in_table = 0;
}

/* Extract document.xml from DOCX and write to temp file */
static int extract_document_xml(const char *docx_path) {
    mz_zip_archive zip;
    memset(&zip, 0, sizeof(zip));
    
    if (!mz_zip_reader_init_file(&zip, docx_path, 0)) {
        return 0;
    }
    
    int file_index = mz_zip_reader_locate_file(&zip, "word/document.xml", NULL, 0);
    if (file_index < 0) {
        mz_zip_reader_end(&zip);
        return 0;
    }
    
    size_t file_size;
    void *file_data = mz_zip_reader_extract_to_heap(&zip, file_index, &file_size, 0);
    if (!file_data) {
        mz_zip_reader_end(&zip);
        return 0;
    }
    
    FILE *temp = fopen(TEMPFILE, "wb");
    if (!temp) {
        mz_free(file_data);
        mz_zip_reader_end(&zip);
        return 0;
    }
    
    fwrite(file_data, 1, file_size, temp);
    fclose(temp);
    
    mz_free(file_data);
    mz_zip_reader_end(&zip);
    return 1;
}

/* Convert DOCX to Markdown */
static void convert_docx_to_md(const char *input_path, const char *output_path) {
    /* Extract document.xml from DOCX */
    if (!extract_document_xml(input_path)) {
        die("Failed to extract document.xml from: %s", input_path);
    }
    
    /* Parse XML */
    struct txml_node *nodes = NULL;
    char *xml_data = txml_parse_file(TEMPFILE, &nodes);
    if (!xml_data) {
        die("Failed to parse document XML");
    }
    
    /* Open output file */
    FILE *output = fopen(output_path, "w");
    if (!output) {
        free(nodes);
        free(xml_data);
        die("Failed to open output file: %s", output_path);
    }
    
    /* Initialize context */
    md_context ctx = {0};
    ctx.output = output;
    
    /* Find document body */
    struct txml_node *body = txml_find(nodes, NULL, TXML_ELEMENT, "w:body", NULL, 1);
    if (!body) {
        fclose(output);
        free(nodes);
        free(xml_data);
        die("No w:body element found in document");
    }
    
    /* Process document content - need to handle paragraphs and tables in order */
    /* Due to txml API, we process them separately which may not preserve exact order */
    struct txml_node *child = body + 1;
    while (child && child->parent >= body && child->type != TXML_EOF) {
        if (child->type == TXML_ELEMENT && child->parent == body) {
            if (strcmp(child->name, "w:p") == 0) {
                process_paragraph(child, &ctx);
            } else if (strcmp(child->name, "w:tbl") == 0) {
                process_table(child, &ctx);
            }
        }
        child++;
    }
    
    /* Cleanup */
    fclose(output);
    free(nodes);
    free(xml_data);
    remove(TEMPFILE);
}

/* Usage information */
static void usage(void) {
    fprintf(stderr, "Usage: docx2md input.docx [-o output.md]\n");
    fprintf(stderr, "Options:\n");
    fprintf(stderr, "  -o FILE    Output file (default: output.md)\n");
    fprintf(stderr, "  -v         Display version information\n");
    fprintf(stderr, "  -h         Display this help message\n");
    exit(1);
}

int main(int argc, char *argv[]) {
    char *input_file = NULL;
    char *output_file = "output.md";
    
    /* Parse arguments */
    for (int i = 1; i < argc; i++) {
        if (strcmp(argv[i], "-v") == 0) {
            printf("docx2md %s\n", VERSION_STR);
            return 0;
        } else if (strcmp(argv[i], "-h") == 0) {
            usage();
        } else if (strcmp(argv[i], "-o") == 0) {
            if (i + 1 >= argc) {
                usage();
            }
            output_file = argv[++i];
        } else if (argv[i][0] == '-') {
            fprintf(stderr, "Unknown option: %s\n", argv[i]);
            usage();
        } else {
            if (input_file) {
                fprintf(stderr, "Multiple input files not allowed\n");
                usage();
            }
            input_file = argv[i];
        }
    }
    
    if (!input_file) {
        fprintf(stderr, "No input file specified\n");
        usage();
    }
    
    convert_docx_to_md(input_file, output_file);
    
    return 0;
}
