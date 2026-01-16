/*
 * docx2md - Convert DOCX to Markdown
 * 
 * Uses txml for XML parsing and miniz for ZIP handling
 */

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <unistd.h>

#include "miniz.h"
#include "util.h"

#define TXML_DEFINE
#include "txml.h"

#define VERSION_STR "0.1"
#define MAX_BUFFER_SIZE (10 * 1024 * 1024)

/* Generate temporary file path - Note: caller should remove the file when done */
static char *get_temp_file_path(void) {
    static char temp_path[256];
    snprintf(temp_path, sizeof(temp_path), "/tmp/docx2md-temp-%d.xml", getpid());
    return temp_path;
}

/* Structure to hold image relationship information */
typedef struct {
    char *rel_id;       /* Relationship ID (e.g., "rId3") */
    char *target;       /* Target path (e.g., "media/image1.png") */
} image_rel;

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
    mz_zip_archive *zip;        /* ZIP archive for extracting images */
    image_rel *image_rels;      /* Array of image relationships */
    int image_rel_count;        /* Number of image relationships */
    const char *output_dir;     /* Directory for output file (for extracting images) */
} md_context;

/* Forward declarations */
static void process_paragraph(struct txml_node *para, md_context *ctx);
static void process_run(struct txml_node *run, md_context *ctx);
static void process_table(struct txml_node *table, md_context *ctx);
static void process_drawing(struct txml_node *drawing, md_context *ctx);
static void parse_relationships(const char *docx_path, md_context *ctx);
static void free_image_rels(md_context *ctx);
static char *xml_unescape(const char *in);

/* XML unescape helper */
static char *xml_unescape(const char *in) {
    if (!in) return NULL;
    
    char *out = malloc(strlen(in) + 1);  // Output will never be longer than input
    if (!out) {
        die("Out of memory in xml_unescape");
    }
    
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

/* Parse relationships from document.xml.rels to get image mappings */
static void parse_relationships(const char *docx_path, md_context *ctx) {
    mz_zip_archive *zip = ctx->zip;
    
    /* Extract document.xml.rels */
    int file_index = mz_zip_reader_locate_file(zip, "word/_rels/document.xml.rels", NULL, 0);
    if (file_index < 0) {
        /* No relationships file - no images */
        return;
    }
    
    size_t file_size;
    void *file_data = mz_zip_reader_extract_to_heap(zip, file_index, &file_size, 0);
    if (!file_data) {
        return;
    }
    
    /* Write to temporary file for txml_parse_file 
     * Note: Using predictable temp path based on PID. For high-security environments,
     * consider using mkstemp() for secure temp file creation. Current approach is
     * acceptable for read-only XML parsing in typical use cases.
     */
    char temp_path[256];
    snprintf(temp_path, sizeof(temp_path), "/tmp/docx2md-rels-%d.xml", getpid());
    FILE *temp = fopen(temp_path, "wb");
    if (!temp) {
        mz_free(file_data);
        return;
    }
    fwrite(file_data, 1, file_size, temp);
    fclose(temp);
    mz_free(file_data);
    
    /* Parse the relationships XML using txml_parse_file */
    struct txml_node *nodes = NULL;
    char *xml_data = txml_parse_file(temp_path, &nodes);
    if (!xml_data) {
        if (nodes) free(nodes);
        remove(temp_path);
        return;
    }
    if (!nodes) {
        free(xml_data);
        remove(temp_path);
        return;
    }
    
    /* Count image relationships first - use recursive search from root */
    ctx->image_rel_count = 0;
    struct txml_node *rel = NULL;
    while ((rel = txml_find(nodes, rel, TXML_ELEMENT, "Relationship", NULL, 1))) {
        struct txml_node *type_attr = txml_find(rel, NULL, TXML_ATTRIBUTE, "Type", NULL, 0);
        if (type_attr && type_attr->value && 
            strstr(type_attr->value, "/image") != NULL) {
            ctx->image_rel_count++;
        }
    }
    
    if (ctx->image_rel_count == 0) {
        free(nodes);
        free(xml_data);
        remove(temp_path);
        return;
    }
    
    /* Allocate array for relationships */
    ctx->image_rels = calloc(ctx->image_rel_count, sizeof(image_rel));
    if (!ctx->image_rels) {
        free(nodes);
        free(xml_data);
        remove(temp_path);
        return;
    }
    
    /* Fill in the relationship data */
    int idx = 0;
    rel = NULL;
    while ((rel = txml_find(nodes, rel, TXML_ELEMENT, "Relationship", NULL, 1)) && idx < ctx->image_rel_count) {
        struct txml_node *type_attr = txml_find(rel, NULL, TXML_ATTRIBUTE, "Type", NULL, 0);
        if (type_attr && type_attr->value && 
            strstr(type_attr->value, "/image") != NULL) {
            
            struct txml_node *id_attr = txml_find(rel, NULL, TXML_ATTRIBUTE, "Id", NULL, 0);
            struct txml_node *target_attr = txml_find(rel, NULL, TXML_ATTRIBUTE, "Target", NULL, 0);
            
            if (id_attr && id_attr->value && target_attr && target_attr->value) {
                char *rel_id = strdup(id_attr->value);
                char *target = strdup(target_attr->value);
                
                if (rel_id && target) {
                    ctx->image_rels[idx].rel_id = rel_id;
                    ctx->image_rels[idx].target = target;
                    idx++;
                } else {
                    /* Handle allocation failure */
                    free(rel_id);
                    free(target);
                }
            }
        }
    }
    
    /* Update actual count to reflect successful allocations */
    ctx->image_rel_count = idx;
    
    free(nodes);
    free(xml_data);
    remove(temp_path);
}

/* Free image relationships */
static void free_image_rels(md_context *ctx) {
    if (ctx->image_rels) {
        for (int i = 0; i < ctx->image_rel_count; i++) {
            free(ctx->image_rels[i].rel_id);
            free(ctx->image_rels[i].target);
        }
        free(ctx->image_rels);
        ctx->image_rels = NULL;
        ctx->image_rel_count = 0;
    }
}

/* Find image target by relationship ID */
static const char *find_image_target(md_context *ctx, const char *rel_id) {
    for (int i = 0; i < ctx->image_rel_count; i++) {
        if (strcmp(ctx->image_rels[i].rel_id, rel_id) == 0) {
            return ctx->image_rels[i].target;
        }
    }
    return NULL;
}

/* Extract image from ZIP to output directory */
static char *extract_image(md_context *ctx, const char *target) {
    /* Build the full path in the ZIP archive - validate length */
    /* Buffer is 512 bytes, need room for "word/" (5 chars) + target + null terminator */
    if (strlen(target) > 506) {
        /* Target path too long */
        return NULL;
    }
    
    char zip_path[512];
    snprintf(zip_path, sizeof(zip_path), "word/%s", target);
    
    /* Find the image in the ZIP */
    int file_index = mz_zip_reader_locate_file(ctx->zip, zip_path, NULL, 0);
    if (file_index < 0) {
        return NULL;
    }
    
    /* Extract to heap */
    size_t file_size;
    void *file_data = mz_zip_reader_extract_to_heap(ctx->zip, file_index, &file_size, 0);
    if (!file_data) {
        return NULL;
    }
    
    /* Get just the filename from the target path */
    const char *filename = strrchr(target, '/');
    if (!filename) {
        filename = target;
    } else {
        filename++; /* Skip the '/' */
    }
    
    /* Build output path */
    char output_path[1024];
    
    /* Validate combined path length */
    /* output_path buffer is 1024 bytes, need room for dir + "/" + filename + null */
    size_t output_dir_len = ctx->output_dir ? strlen(ctx->output_dir) : 0;
    size_t filename_len = strlen(filename);
    if (output_dir_len + filename_len + 2 > sizeof(output_path) - 1) {
        /* Combined path too long */
        mz_free(file_data);
        return NULL;
    }
    
    /* Build the actual output path */
    if (ctx->output_dir && ctx->output_dir[0] != '\0') {
        snprintf(output_path, sizeof(output_path), "%s/%s", ctx->output_dir, filename);
    } else {
        snprintf(output_path, sizeof(output_path), "%s", filename);
    }
    
    /* Write the image file */
    FILE *img_file = fopen(output_path, "wb");
    if (!img_file) {
        mz_free(file_data);
        return NULL;
    }
    
    fwrite(file_data, 1, file_size, img_file);
    fclose(img_file);
    mz_free(file_data);
    
    /* Return just the filename for markdown */
    char *result = strdup(filename);
    if (!result) {
        /* Allocation failure - remove the written file to maintain consistency */
        remove(output_path);
    }
    return result;
}

/* Process a drawing element (image) */
static void process_drawing(struct txml_node *drawing, md_context *ctx) {
    /* Find the blip element which contains the relationship ID */
    struct txml_node *blip = txml_find(drawing, NULL, TXML_ELEMENT, "a:blip", NULL, 1);
    if (!blip) {
        return;
    }
    
    /* Get the r:embed attribute */
    struct txml_node *embed_attr = txml_find(blip, NULL, TXML_ATTRIBUTE, "r:embed", NULL, 0);
    if (!embed_attr || !embed_attr->value) {
        return;
    }
    
    /* Find the image target path using the relationship ID */
    const char *target = find_image_target(ctx, embed_attr->value);
    if (!target) {
        return;
    }
    
    /* Extract the image to the output directory */
    char *image_filename = extract_image(ctx, target);
    if (!image_filename) {
        return;
    }
    
    /* Find alt text from docPr name attribute */
    struct txml_node *docPr = txml_find(drawing, NULL, TXML_ELEMENT, "wp:docPr", NULL, 1);
    const char *alt_text = "Image";
    if (docPr) {
        struct txml_node *name_attr = txml_find(docPr, NULL, TXML_ATTRIBUTE, "name", NULL, 0);
        if (name_attr && name_attr->value) {
            alt_text = name_attr->value;
        }
    }
    
    /* Output markdown image syntax */
    fprintf(ctx->output, "![%s](%s)", alt_text, image_filename);
    
    free(image_filename);
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
    
    /* Check for drawing (image) first */
    struct txml_node *drawing = txml_find(run, NULL, TXML_ELEMENT, "w:drawing", NULL, 0);
    if (drawing) {
        process_drawing(drawing, ctx);
        /* Reset to old state */
        ctx->in_bold = old_bold;
        ctx->in_italic = old_italic;
        ctx->in_code = old_code;
        ctx->in_strikethrough = old_strike;
        return;
    }
    
    /* Extract text content first to check if run is empty */
    struct txml_node *text_node = txml_find(run, NULL, TXML_ELEMENT, "w:t", NULL, 0);
    char *unescaped = NULL;
    int has_text = 0;
    
    if (text_node && text_node->value) {
        unescaped = xml_unescape(text_node->value);
        if (unescaped && unescaped[0] != '\0') {
            has_text = 1;
        }
    }
    
    /* Check for line break */
    struct txml_node *br = txml_find(run, NULL, TXML_ELEMENT, "w:br", NULL, 0);
    
    /* Skip empty runs (no text and no line break), but reset to old state */
    if (!has_text && !br) {
        if (unescaped) free(unescaped);
        /* Reset formatting to old state since we're skipping this run */
        ctx->in_bold = old_bold;
        ctx->in_italic = old_italic;
        ctx->in_code = old_code;
        ctx->in_strikethrough = old_strike;
        return;
    }
    
    /* Open formatting markers only if run has content */
    if (ctx->in_strikethrough && !old_strike) fprintf(ctx->output, "~~");
    if (ctx->in_bold && !old_bold) fprintf(ctx->output, "**");
    if (ctx->in_italic && !old_italic) fprintf(ctx->output, "*");
    if (ctx->in_code && !old_code) fprintf(ctx->output, "`");
    
    /* Output text content */
    if (has_text && unescaped) {
        fprintf(ctx->output, "%s", unescaped);
        free(unescaped);
    }
    
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

/* Convert DOCX to Markdown */
static void convert_docx_to_md(const char *input_path, const char *output_path) {
    const char *temp_file = get_temp_file_path();
    
    /* Open ZIP archive for image extraction */
    mz_zip_archive zip;
    memset(&zip, 0, sizeof(zip));
    
    if (!mz_zip_reader_init_file(&zip, input_path, 0)) {
        die("Failed to open DOCX file: %s", input_path);
    }
    
    /* Extract document.xml from DOCX */
    int file_index = mz_zip_reader_locate_file(&zip, "word/document.xml", NULL, 0);
    if (file_index < 0) {
        mz_zip_reader_end(&zip);
        die("Failed to find document.xml in: %s", input_path);
    }
    
    size_t file_size;
    void *file_data = mz_zip_reader_extract_to_heap(&zip, file_index, &file_size, 0);
    if (!file_data) {
        mz_zip_reader_end(&zip);
        die("Failed to extract document.xml from: %s", input_path);
    }
    
    FILE *temp = fopen(temp_file, "wb");
    if (!temp) {
        mz_free(file_data);
        mz_zip_reader_end(&zip);
        die("Failed to create temporary file: %s", temp_file);
    }
    
    fwrite(file_data, 1, file_size, temp);
    fclose(temp);
    mz_free(file_data);
    
    /* Parse XML */
    struct txml_node *nodes = NULL;
    char *xml_data = txml_parse_file((char *)temp_file, &nodes);  /* txml_parse_file modifies path string */
    if (!xml_data) {
        mz_zip_reader_end(&zip);
        die("Failed to parse document XML");
    }
    
    /* Open output file */
    FILE *output = fopen(output_path, "w");
    if (!output) {
        free(nodes);
        free(xml_data);
        mz_zip_reader_end(&zip);
        die("Failed to open output file: %s", output_path);
    }
    
    /* Extract output directory from output path */
    char output_dir[1024] = "";
    const char *last_slash = strrchr(output_path, '/');
    if (last_slash) {
        size_t dir_len = last_slash - output_path;
        if (dir_len > 0 && dir_len < sizeof(output_dir) - 1) {
            memcpy(output_dir, output_path, dir_len);
            output_dir[dir_len] = '\0';
        }
    }
    
    /* Initialize context */
    md_context ctx = {0};
    ctx.output = output;
    ctx.zip = &zip;
    ctx.output_dir = output_dir[0] ? output_dir : NULL;
    
    /* Parse relationships to find image mappings */
    parse_relationships(input_path, &ctx);
    
    /* Find document body */
    struct txml_node *body = txml_find(nodes, NULL, TXML_ELEMENT, "w:body", NULL, 1);
    if (!body) {
        fclose(output);
        free(nodes);
        free(xml_data);
        free_image_rels(&ctx);
        mz_zip_reader_end(&zip);
        die("No w:body element found in document");
    }
    
    /* Process document content in order - paragraphs and tables 
     * Note: We use pointer arithmetic here because txml stores nodes in a contiguous array.
     * This is the documented way to traverse nodes in document order.
     * Using txml_find() would not preserve the interleaved order of paragraphs and tables.
     */
    struct txml_node *child = body + 1;
    while (child && child->type != TXML_EOF) {
        /* Check if this node is a direct child of body */
        if (child->parent == body && child->type == TXML_ELEMENT) {
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
    free_image_rels(&ctx);
    mz_zip_reader_end(&zip);
    remove(temp_file);
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
