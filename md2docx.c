/*
 * md2docx - Convert Markdown to DOCX
 * 
 * Uses md4c for Markdown parsing and miniz for ZIP handling
 */

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <sys/stat.h>
#include <errno.h>

#include "md4c.h"
#include "miniz.h"
#include "util.h"

#define VERSION_STR "0.1"
#define MAX_BUFFER_SIZE (10 * 1024 * 1024)  // 10MB buffer for document

/* Context structure to hold state during parsing */
typedef struct {
    char *xml_buffer;
    size_t xml_size;
    size_t xml_capacity;
    int list_level;
    int in_paragraph;
    int in_list_item;
    int para_has_content;
    int in_span;
    int in_text_run;
    char **image_paths;
    int image_count;
    int image_capacity;
    int next_image_id;
} docx_context;

/* Forward declarations */
static void append_xml(docx_context *ctx, const char *str);
static void append_xml_n(docx_context *ctx, const char *str, size_t len);
static void xml_escape_append(docx_context *ctx, const char *text, size_t size);
static void ensure_paragraph(docx_context *ctx);
static void close_paragraph(docx_context *ctx);

/* XML escaping */
static void xml_escape_append(docx_context *ctx, const char *text, size_t size)
{
    for (size_t i = 0; i < size; i++) {
        switch (text[i]) {
            case '&':  append_xml(ctx, "&amp;"); break;
            case '<':  append_xml(ctx, "&lt;"); break;
            case '>':  append_xml(ctx, "&gt;"); break;
            case '"':  append_xml(ctx, "&quot;"); break;
            case '\'': append_xml(ctx, "&apos;"); break;
            default:   
                if (ctx->xml_size >= ctx->xml_capacity - 1) {
                    ctx->xml_capacity *= 2;
                    char *new_buffer = realloc(ctx->xml_buffer, ctx->xml_capacity);
                    if (!new_buffer) die("Out of memory");
                    ctx->xml_buffer = new_buffer;
                }
                ctx->xml_buffer[ctx->xml_size++] = text[i];
                break;
        }
    }
}

/* Append string to XML buffer */
static void append_xml(docx_context *ctx, const char *str)
{
    size_t len = strlen(str);
    while (ctx->xml_size + len >= ctx->xml_capacity) {
        ctx->xml_capacity *= 2;
        char *new_buffer = realloc(ctx->xml_buffer, ctx->xml_capacity);
        if (!new_buffer) die("Out of memory");
        ctx->xml_buffer = new_buffer;
    }
    memcpy(ctx->xml_buffer + ctx->xml_size, str, len);
    ctx->xml_size += len;
}

static void append_xml_n(docx_context *ctx, const char *str, size_t len)
{
    while (ctx->xml_size + len >= ctx->xml_capacity) {
        ctx->xml_capacity *= 2;
        char *new_buffer = realloc(ctx->xml_buffer, ctx->xml_capacity);
        if (!new_buffer) die("Out of memory");
        ctx->xml_buffer = new_buffer;
    }
    memcpy(ctx->xml_buffer + ctx->xml_size, str, len);
    ctx->xml_size += len;
}

/* Ensure we're in a paragraph */
static void ensure_paragraph(docx_context *ctx)
{
    if (!ctx->in_paragraph) {
        append_xml(ctx, "<w:p>");
        ctx->in_paragraph = 1;
        ctx->para_has_content = 0;
    }
}

static void close_paragraph(docx_context *ctx)
{
    if (ctx->in_paragraph) {
        // Close any open text run
        if (ctx->in_text_run) {
            append_xml(ctx, "</w:t></w:r>");
            ctx->in_text_run = 0;
        }
        append_xml(ctx, "</w:p>");
        ctx->in_paragraph = 0;
        ctx->para_has_content = 0;
        ctx->in_span = 0;
    }
}

/* Add image to tracking list */
static void add_image(docx_context *ctx, const char *path, size_t path_len)
{
    if (ctx->image_count >= ctx->image_capacity) {
        ctx->image_capacity = ctx->image_capacity ? ctx->image_capacity * 2 : 8;
        char **new_paths = realloc(ctx->image_paths, ctx->image_capacity * sizeof(char*));
        if (!new_paths) die("Out of memory");
        ctx->image_paths = new_paths;
    }
    
    ctx->image_paths[ctx->image_count] = malloc(path_len + 1);
    if (!ctx->image_paths[ctx->image_count]) die("Out of memory");
    memcpy(ctx->image_paths[ctx->image_count], path, path_len);
    ctx->image_paths[ctx->image_count][path_len] = '\0';
    ctx->image_count++;
}

/* Free image paths */
static void free_image_paths(docx_context *ctx)
{
    if (ctx->image_paths) {
        for (int i = 0; i < ctx->image_count; i++) {
            free(ctx->image_paths[i]);
        }
        free(ctx->image_paths);
        ctx->image_paths = NULL;
        ctx->image_count = 0;
    }
}

/* Markdown callback: enter block */
static int enter_block_callback(MD_BLOCKTYPE type, void *detail, void *userdata)
{
    docx_context *ctx = (docx_context *)userdata;
    
    switch (type) {
        case MD_BLOCK_DOC:
            // Document start - already handled
            break;
            
        case MD_BLOCK_QUOTE:
            ensure_paragraph(ctx);
            break;
            
        case MD_BLOCK_UL:
        case MD_BLOCK_OL:
            close_paragraph(ctx);
            ctx->list_level++;
            break;
            
        case MD_BLOCK_LI:
            close_paragraph(ctx);
            ctx->in_list_item = 1;
            break;
            
        case MD_BLOCK_HR:
            close_paragraph(ctx);
            append_xml(ctx, "<w:p><w:pPr><w:pBdr><w:bottom w:val=\"single\" w:sz=\"6\" w:space=\"1\" w:color=\"auto\"/></w:pBdr></w:pPr></w:p>");
            break;
            
        case MD_BLOCK_H: {
            close_paragraph(ctx);
            MD_BLOCK_H_DETAIL *h = (MD_BLOCK_H_DETAIL *)detail;
            char buf[256];
            snprintf(buf, sizeof(buf), "<w:p><w:pPr><w:pStyle w:val=\"Heading%u\"/></w:pPr>", h->level);
            append_xml(ctx, buf);
            ctx->in_paragraph = 1;
            ctx->para_has_content = 0;
            break;
        }
            
        case MD_BLOCK_CODE: {
            close_paragraph(ctx);
            append_xml(ctx, "<w:p><w:pPr><w:pStyle w:val=\"Code\"/></w:pPr>");
            ctx->in_paragraph = 1;
            ctx->para_has_content = 0;
            break;
        }
            
        case MD_BLOCK_HTML:
            // Skip raw HTML
            break;
            
        case MD_BLOCK_P:
            close_paragraph(ctx);
            if (ctx->in_list_item) {
                append_xml(ctx, "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>");
            } else {
                append_xml(ctx, "<w:p>");
            }
            ctx->in_paragraph = 1;
            break;
            
        case MD_BLOCK_TABLE:
            close_paragraph(ctx);
            append_xml(ctx, "<w:tbl><w:tblPr><w:tblStyle w:val=\"TableGrid\"/><w:tblW w:w=\"5000\" w:type=\"pct\"/></w:tblPr>");
            break;
            
        case MD_BLOCK_THEAD:
        case MD_BLOCK_TBODY:
            // Just grouping, no special output
            break;
            
        case MD_BLOCK_TR:
            append_xml(ctx, "<w:tr>");
            break;
            
        case MD_BLOCK_TH:
        case MD_BLOCK_TD:
            append_xml(ctx, "<w:tc><w:tcPr><w:tcW w:w=\"0\" w:type=\"auto\"/></w:tcPr><w:p>");
            ctx->in_paragraph = 1;
            break;
            
        default:
            break;
    }
    
    return 0;
}

/* Markdown callback: leave block */
static int leave_block_callback(MD_BLOCKTYPE type, void *detail, void *userdata)
{
    docx_context *ctx = (docx_context *)userdata;
    
    switch (type) {
        case MD_BLOCK_DOC:
            close_paragraph(ctx);
            break;
            
        case MD_BLOCK_QUOTE:
            break;
            
        case MD_BLOCK_UL:
        case MD_BLOCK_OL:
            ctx->list_level--;
            break;
            
        case MD_BLOCK_LI:
            ctx->in_list_item = 0;
            break;
            
        case MD_BLOCK_HR:
            break;
            
        case MD_BLOCK_H:
            close_paragraph(ctx);
            break;
            
        case MD_BLOCK_CODE:
            close_paragraph(ctx);
            break;
            
        case MD_BLOCK_HTML:
            break;
            
        case MD_BLOCK_P:
            close_paragraph(ctx);
            break;
            
        case MD_BLOCK_TABLE:
            append_xml(ctx, "</w:tbl>");
            break;
            
        case MD_BLOCK_THEAD:
        case MD_BLOCK_TBODY:
            break;
            
        case MD_BLOCK_TR:
            append_xml(ctx, "</w:tr>");
            break;
            
        case MD_BLOCK_TH:
        case MD_BLOCK_TD:
            close_paragraph(ctx);
            append_xml(ctx, "</w:tc>");
            break;
            
        default:
            break;
    }
    
    return 0;
}

/* Markdown callback: enter span */
static int enter_span_callback(MD_SPANTYPE type, void *detail, void *userdata)
{
    docx_context *ctx = (docx_context *)userdata;
    
    ensure_paragraph(ctx);
    
    // Close any open text run before starting a new styled one
    if (ctx->in_text_run) {
        append_xml(ctx, "</w:t></w:r>");
        ctx->in_text_run = 0;
    }
    
    switch (type) {
        case MD_SPAN_EM:
            append_xml(ctx, "<w:r><w:rPr><w:i/></w:rPr><w:t>");
            ctx->in_span = 1;
            break;
            
        case MD_SPAN_STRONG:
            append_xml(ctx, "<w:r><w:rPr><w:b/></w:rPr><w:t>");
            ctx->in_span = 1;
            break;
            
        case MD_SPAN_A: {
            // For now, just render as underlined text with href in []
            append_xml(ctx, "<w:r><w:rPr><w:u w:val=\"single\"/></w:rPr><w:t>");
            ctx->in_span = 1;
            break;
        }
            
        case MD_SPAN_IMG: {
            MD_SPAN_IMG_DETAIL *img = (MD_SPAN_IMG_DETAIL *)detail;
            // Track image for embedding
            if (img->src.size > 0) {
                int current_img_index = ctx->image_count;
                add_image(ctx, img->src.text, img->src.size);
                
                // Embed image in document
                // Image relationship IDs start at rId3 (rId1=styles, rId2=numbering)
                int rel_id = current_img_index + 3;
                char buf[1024];
                int img_id = ctx->next_image_id++;
                snprintf(buf, sizeof(buf),
                    "<w:r><w:drawing><wp:inline><wp:extent cx=\"2000000\" cy=\"2000000\"/>"
                    "<wp:docPr id=\"%d\" name=\"Image%d\"/>"
                    "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
                    "<pic:pic><pic:nvPicPr><pic:cNvPr id=\"%d\" name=\"Image%d\"/>"
                    "<pic:cNvPicPr/></pic:nvPicPr>"
                    "<pic:blipFill><a:blip r:embed=\"rId%d\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>"
                    "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"2000000\" cy=\"2000000\"/></a:xfrm>"
                    "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr>"
                    "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>",
                    img_id, img_id, img_id, img_id, rel_id);
                append_xml(ctx, buf);
            }
            // Images don't have content text, so don't set in_span
            break;
        }
            
        case MD_SPAN_CODE:
            append_xml(ctx, "<w:r><w:rPr><w:rStyle w:val=\"CodeChar\"/></w:rPr><w:t xml:space=\"preserve\">");
            ctx->in_span = 1;
            break;
            
        case MD_SPAN_DEL:
            append_xml(ctx, "<w:r><w:rPr><w:strike/></w:rPr><w:t>");
            ctx->in_span = 1;
            break;
            
        case MD_SPAN_U:
            append_xml(ctx, "<w:r><w:rPr><w:u w:val=\"single\"/></w:rPr><w:t>");
            ctx->in_span = 1;
            break;
            
        default:
            ctx->in_span = 1;
            break;
    }
    
    ctx->para_has_content = 1;
    return 0;
}

/* Markdown callback: leave span */
static int leave_span_callback(MD_SPANTYPE type, void *detail, void *userdata)
{
    docx_context *ctx = (docx_context *)userdata;
    
    switch (type) {
        case MD_SPAN_EM:
        case MD_SPAN_STRONG:
        case MD_SPAN_A:
        case MD_SPAN_CODE:
        case MD_SPAN_DEL:
        case MD_SPAN_U:
            append_xml(ctx, "</w:t></w:r>");
            ctx->in_span = 0;
            break;
            
        case MD_SPAN_IMG:
            // Image already embedded, no text to close
            break;
            
        default:
            ctx->in_span = 0;
            break;
    }
    
    return 0;
}

/* Markdown callback: text */
static int text_callback(MD_TEXTTYPE type, const MD_CHAR *text, MD_SIZE size, void *userdata)
{
    docx_context *ctx = (docx_context *)userdata;
    
    ensure_paragraph(ctx);
    
    switch (type) {
        case MD_TEXT_NORMAL:
        case MD_TEXT_CODE:
        case MD_TEXT_HTML:
            // If we're in a span, text is already wrapped properly
            // Otherwise, we need to create a text run
            if (!ctx->in_span && !ctx->in_text_run) {
                append_xml(ctx, "<w:r><w:t>");
                ctx->in_text_run = 1;
                ctx->para_has_content = 1;
            }
            xml_escape_append(ctx, text, size);
            break;
            
        case MD_TEXT_NULLCHAR:
            break;
            
        case MD_TEXT_BR:
        case MD_TEXT_SOFTBR:
            if (ctx->in_text_run) {
                append_xml(ctx, "</w:t></w:r>");
                ctx->in_text_run = 0;
            }
            append_xml(ctx, "<w:r><w:br/></w:r>");
            break;
            
        case MD_TEXT_ENTITY:
            // Decode common entities
            if (size == 4 && strncmp(text, "&lt;", 4) == 0) {
                xml_escape_append(ctx, "<", 1);
            } else if (size == 4 && strncmp(text, "&gt;", 4) == 0) {
                xml_escape_append(ctx, ">", 1);
            } else if (size == 5 && strncmp(text, "&amp;", 5) == 0) {
                xml_escape_append(ctx, "&", 1);
            } else if (size == 6 && strncmp(text, "&quot;", 6) == 0) {
                xml_escape_append(ctx, "\"", 1);
            } else if (size == 6 && strncmp(text, "&apos;", 6) == 0) {
                xml_escape_append(ctx, "'", 1);
            } else {
                // Other entities - just output as-is
                xml_escape_append(ctx, text, size);
            }
            break;
            
        default:
            if (!ctx->in_span && !ctx->in_text_run) {
                append_xml(ctx, "<w:r><w:t>");
                ctx->in_text_run = 1;
            }
            xml_escape_append(ctx, text, size);
            break;
    }
    
    return 0;
}

/* Read file into memory */
static char *read_file(const char *filename, size_t *size)
{
    FILE *f = fopen(filename, "rb");
    if (!f) {
        return NULL;
    }
    
    fseek(f, 0, SEEK_END);
    *size = ftell(f);
    fseek(f, 0, SEEK_SET);
    
    char *buffer = malloc(*size + 1);
    if (!buffer) {
        fclose(f);
        return NULL;
    }
    
    if (fread(buffer, 1, *size, f) != *size) {
        free(buffer);
        fclose(f);
        return NULL;
    }
    
    buffer[*size] = '\0';
    fclose(f);
    return buffer;
}

/* Create the [Content_Types].xml file */
static const char *get_content_types_xml(void)
{
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
           "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
           "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
           "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
           "<Default Extension=\"png\" ContentType=\"image/png\"/>"
           "<Default Extension=\"jpg\" ContentType=\"image/jpeg\"/>"
           "<Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>"
           "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
           "<Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>"
           "<Override PartName=\"/word/numbering.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml\"/>"
           "</Types>";
}

/* Create the _rels/.rels file */
static const char *get_rels_xml(void)
{
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
           "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
           "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
           "</Relationships>";
}

/* Create the word/_rels/document.xml.rels file */
static char *get_document_rels_xml(docx_context *ctx)
{
    size_t capacity = 4096;
    char *xml = malloc(capacity);
    if (!xml) return NULL;
    
    size_t len = snprintf(xml, capacity,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering\" Target=\"numbering.xml\"/>");
    
    // Add image relationships
    for (int i = 0; i < ctx->image_count; i++) {
        char buf[512];
        const char *ext = strrchr(ctx->image_paths[i], '.');
        if (!ext) ext = ".png";
        
        int n = snprintf(buf, sizeof(buf),
            "<Relationship Id=\"rId%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image%d%s\"/>",
            i + 3, i + 1, ext);
        
        while (len + n >= capacity) {
            capacity *= 2;
            char *new_xml = realloc(xml, capacity);
            if (!new_xml) {
                free(xml);
                return NULL;
            }
            xml = new_xml;
        }
        
        memcpy(xml + len, buf, n);
        len += n;
    }
    
    const char *end = "</Relationships>";
    while (len + strlen(end) >= capacity) {
        capacity *= 2;
        char *new_xml = realloc(xml, capacity);
        if (!new_xml) {
            free(xml);
            return NULL;
        }
        xml = new_xml;
    }
    strcpy(xml + len, end);
    
    return xml;
}

/* Create the word/styles.xml file */
static const char *get_styles_xml(void)
{
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
           "<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
           "<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\" w:cs=\"Calibri\"/><w:sz w:val=\"22\"/></w:rPr></w:rPrDefault></w:docDefaults>"
           "<w:style w:type=\"paragraph\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/><w:qFormat/></w:style>"
           "<w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"Heading 1\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:keepNext/><w:spacing w:before=\"480\" w:after=\"0\"/></w:pPr><w:rPr><w:b/><w:sz w:val=\"32\"/></w:rPr></w:style>"
           "<w:style w:type=\"paragraph\" w:styleId=\"Heading2\"><w:name w:val=\"Heading 2\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:keepNext/><w:spacing w:before=\"200\" w:after=\"0\"/></w:pPr><w:rPr><w:b/><w:sz w:val=\"28\"/></w:rPr></w:style>"
           "<w:style w:type=\"paragraph\" w:styleId=\"Heading3\"><w:name w:val=\"Heading 3\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:keepNext/><w:spacing w:before=\"200\" w:after=\"0\"/></w:pPr><w:rPr><w:b/><w:sz w:val=\"26\"/></w:rPr></w:style>"
           "<w:style w:type=\"paragraph\" w:styleId=\"Heading4\"><w:name w:val=\"Heading 4\"/><w:basedOn w:val=\"Normal\"/><w:rPr><w:b/><w:sz w:val=\"24\"/></w:rPr></w:style>"
           "<w:style w:type=\"paragraph\" w:styleId=\"Heading5\"><w:name w:val=\"Heading 5\"/><w:basedOn w:val=\"Normal\"/><w:rPr><w:b/></w:rPr></w:style>"
           "<w:style w:type=\"paragraph\" w:styleId=\"Heading6\"><w:name w:val=\"Heading 6\"/><w:basedOn w:val=\"Normal\"/><w:rPr><w:b/><w:i/></w:rPr></w:style>"
           "<w:style w:type=\"paragraph\" w:styleId=\"Code\"><w:name w:val=\"Code\"/><w:basedOn w:val=\"Normal\"/><w:rPr><w:rFonts w:ascii=\"Courier New\" w:hAnsi=\"Courier New\"/><w:sz w:val=\"20\"/></w:rPr></w:style>"
           "<w:style w:type=\"character\" w:styleId=\"CodeChar\"><w:name w:val=\"Code Char\"/><w:rPr><w:rFonts w:ascii=\"Courier New\" w:hAnsi=\"Courier New\"/></w:rPr></w:style>"
           "</w:styles>";
}

/* Create the word/numbering.xml file for lists */
static const char *get_numbering_xml(void)
{
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
           "<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
           "<w:abstractNum w:abstractNumId=\"0\">"
           "<w:multiLevelType w:val=\"hybridMultilevel\"/>"
           "<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"\"/><w:lvlJc w:val=\"left\"/></w:lvl>"
           "</w:abstractNum>"
           "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>"
           "</w:numbering>";
}

/* Create the main document.xml with content */
static char *get_document_xml(docx_context *ctx)
{
    size_t capacity = ctx->xml_size + 1024;
    char *xml = malloc(capacity);
    if (!xml) return NULL;
    
    const char *header = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
                        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
                        "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
                        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
                        "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
                        "<w:body>";
    const char *footer = "</w:body></w:document>";
    
    size_t header_len = strlen(header);
    size_t footer_len = strlen(footer);
    size_t total = header_len + ctx->xml_size + footer_len + 1;
    
    if (total > capacity) {
        capacity = total;
        char *new_xml = realloc(xml, capacity);
        if (!new_xml) {
            free(xml);
            return NULL;
        }
        xml = new_xml;
    }
    
    strcpy(xml, header);
    memcpy(xml + header_len, ctx->xml_buffer, ctx->xml_size);
    strcpy(xml + header_len + ctx->xml_size, footer);
    
    return xml;
}

/* Add file to ZIP archive */
static int add_file_to_zip(mz_zip_archive *zip, const char *archive_name, const void *data, size_t size)
{
    return mz_zip_writer_add_mem(zip, archive_name, data, size, MZ_DEFAULT_COMPRESSION);
}

/* Convert markdown to DOCX */
static int convert_markdown_to_docx(const char *md_file, const char *docx_file)
{
    size_t md_size;
    char *md_content = read_file(md_file, &md_size);
    if (!md_content) {
        fprintf(stderr, "Error: Cannot read input file '%s'\n", md_file);
        return 1;
    }
    
    // Initialize context
    docx_context ctx = {0};
    ctx.xml_capacity = 64 * 1024;
    ctx.xml_buffer = malloc(ctx.xml_capacity);
    if (!ctx.xml_buffer) {
        free(md_content);
        die("Out of memory");
    }
    ctx.xml_size = 0;
    ctx.next_image_id = 1;
    
    // Set up MD4C parser with GitHub-flavored markdown
    MD_PARSER parser = {0};
    parser.abi_version = 0;
    parser.flags = MD_FLAG_TABLES | MD_FLAG_STRIKETHROUGH | MD_FLAG_TASKLISTS | 
                   MD_FLAG_PERMISSIVEAUTOLINKS | MD_FLAG_UNDERLINE;
    parser.enter_block = enter_block_callback;
    parser.leave_block = leave_block_callback;
    parser.enter_span = enter_span_callback;
    parser.leave_span = leave_span_callback;
    parser.text = text_callback;
    
    // Parse markdown
    int ret = md_parse(md_content, md_size, &parser, &ctx);
    free(md_content);
    
    if (ret != 0) {
        fprintf(stderr, "Error: Failed to parse markdown (code %d)\n", ret);
        free(ctx.xml_buffer);
        free_image_paths(&ctx);
        return 1;
    }
    
    // Null-terminate XML buffer
    if (ctx.xml_size >= ctx.xml_capacity) {
        char *new_buffer = realloc(ctx.xml_buffer, ctx.xml_size + 1);
        if (!new_buffer) {
            free(ctx.xml_buffer);
            free_image_paths(&ctx);
            die("Out of memory");
        }
        ctx.xml_buffer = new_buffer;
    }
    ctx.xml_buffer[ctx.xml_size] = '\0';
    
    // Create DOCX (ZIP archive)
    mz_zip_archive zip = {0};
    if (!mz_zip_writer_init_file(&zip, docx_file, 0)) {
        fprintf(stderr, "Error: Cannot create output file '%s'\n", docx_file);
        free(ctx.xml_buffer);
        free_image_paths(&ctx);
        return 1;
    }
    
    // Add required files to ZIP
    const char *content_types = get_content_types_xml();
    add_file_to_zip(&zip, "[Content_Types].xml", content_types, strlen(content_types));
    
    const char *rels = get_rels_xml();
    add_file_to_zip(&zip, "_rels/.rels", rels, strlen(rels));
    
    char *doc_rels = get_document_rels_xml(&ctx);
    if (doc_rels) {
        add_file_to_zip(&zip, "word/_rels/document.xml.rels", doc_rels, strlen(doc_rels));
        free(doc_rels);
    }
    
    char *document = get_document_xml(&ctx);
    if (document) {
        add_file_to_zip(&zip, "word/document.xml", document, strlen(document));
        free(document);
    }
    
    const char *styles = get_styles_xml();
    add_file_to_zip(&zip, "word/styles.xml", styles, strlen(styles));
    
    const char *numbering = get_numbering_xml();
    add_file_to_zip(&zip, "word/numbering.xml", numbering, strlen(numbering));
    
    // Add images if any
    for (int i = 0; i < ctx.image_count; i++) {
        size_t img_size;
        char *img_data = read_file(ctx.image_paths[i], &img_size);
        if (img_data) {
            char archive_name[256];
            const char *ext = strrchr(ctx.image_paths[i], '.');
            if (!ext) ext = ".png";
            snprintf(archive_name, sizeof(archive_name), "word/media/image%d%s", i + 1, ext);
            add_file_to_zip(&zip, archive_name, img_data, img_size);
            free(img_data);
        }
    }
    
    // Finalize ZIP
    if (!mz_zip_writer_finalize_archive(&zip)) {
        fprintf(stderr, "Error: Failed to finalize ZIP archive\n");
        mz_zip_writer_end(&zip);
        free(ctx.xml_buffer);
        free_image_paths(&ctx);
        return 1;
    }
    
    mz_zip_writer_end(&zip);
    free(ctx.xml_buffer);
    free_image_paths(&ctx);
    
    printf("Successfully converted '%s' to '%s'\n", md_file, docx_file);
    return 0;
}

static void usage(void)
{
    fprintf(stderr, "usage: md2docx input.md [-o output.docx]\n");
    fprintf(stderr, "\nOptions:\n");
    fprintf(stderr, "  -o FILE    Specify output file (default: output.docx)\n");
    fprintf(stderr, "  -v         Display version information\n");
    fprintf(stderr, "  -h         Display this help message\n");
    exit(1);
}

int main(int argc, char *argv[])
{
    char *input_file = NULL;
    char *output_file = "output.docx";
    
    // Parse arguments
    for (int i = 1; i < argc; i++) {
        if (strcmp(argv[i], "-v") == 0) {
            printf("md2docx version %s\n", VERSION_STR);
            return 0;
        } else if (strcmp(argv[i], "-h") == 0) {
            usage();
        } else if (strcmp(argv[i], "-o") == 0) {
            if (i + 1 >= argc) {
                fprintf(stderr, "Error: -o requires an argument\n");
                usage();
            }
            output_file = argv[++i];
        } else if (argv[i][0] == '-') {
            fprintf(stderr, "Error: Unknown option '%s'\n", argv[i]);
            usage();
        } else {
            if (input_file) {
                fprintf(stderr, "Error: Multiple input files specified\n");
                usage();
            }
            input_file = argv[i];
        }
    }
    
    if (!input_file) {
        fprintf(stderr, "Error: No input file specified\n");
        usage();
    }
    
    return convert_markdown_to_docx(input_file, output_file);
}
