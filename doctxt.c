// // Compile: gcc ycml.c cml2tsv.c
// #include <stdio.h>
// #include <stdlib.h>
// #include <string.h>
// #include <assert.h>

// #define TXML_DEFINE
// #include "txml.h"

// #define count(ARRAY) (sizeof (ARRAY) / sizeof *(ARRAY))

// #define BUFSIZE 4096

// char *lread_file(const char *filename)
// {
// 	char *buffer = 0;
// 	size_t length;
	
// 	// Open file
// 	FILE *file = fopen(filename, "rb");
// 	if (!file) {
// 		fprintf(stderr, "Unanble to open %s\n", filename);
// 		return NULL;
// 	}

// 	// Determine the end of the file
// 	fseek(file, 0, SEEK_END);
// 	// Get the length
// 	length = ftell(file);
// 	// Rewind file
// 	fseek(file, 0, SEEK_SET);

// 	// Allocate space
// 	buffer = malloc(length + 1);
// 	// Read file content
// 	fread(buffer, 1, length, file);

// 	// Close the file
// 	fclose(file);

// 	// Terminate with \0 !
// 	buffer[length] = '\0';

// 	return buffer;
// }

// #define NR_NODES (1024*1024*1024)

// int main(int argc, char *argv[])
// {
// 	if (argc != 2) {
// 		fprintf(stderr, "usage: %s uniprot.xml\n", argv[0]);
// 		return 1;
// 	}

// 	char *doc = lread_file(argv[1]); /* The XML document as a zero-terminated string */
// 	if (!doc) {
// 		return 1;
// 	}
// 	struct txml_node *nodes = malloc(sizeof(struct txml_node) * NR_NODES);
// 	if (!nodes) {
// 		fprintf(stderr, "Failed to allocate memory\n");
// 		return 1;
// 	}
// 	char *tail = txml_parse(doc, NR_NODES, nodes);
// 	assert(!tail); // check for parsing success

// 	puts("--- tx ---"); // use tx API
// 	const char *path[] = { "feature", "location", "begin", "@position" };
// 	const char *path2[] = { "@type" };
// 	const char *path3[] = { "end", "@position" };
// 	struct txml_node *current_bar = NULL;
// 	struct txml_node *current_bar2 = NULL;
// 	struct txml_node *current_bar3 = NULL;
// 	while((current_bar = txml_get(nodes, current_bar, count(path), path))) {
// 		while((current_bar2 = txml_get(current_bar->parent->parent->parent, current_bar2, count(path2), path2))) {
// 			printf("%s\t", current_bar2->value); // type
// 		}

// 		//puts(current_bar->name);
// 		printf("%s\t", current_bar->value); // begin
// 		// puts(current_bar->parent->name);
// 		//puts(current_bar->parent->parent->name);
// 		// puts(current_bar->parent->parent->parent->name);
// 		// puts(current_bar->parent->parent->parent->parent->name);

// 		while((current_bar3 = txml_get(current_bar->parent->parent, current_bar3, count(path3), path3))) {
// 			puts(current_bar3->value); // end
// 		}
// 		//printf("%s %s %s\n", current_bar2->value, current_bar->value, current_bar3->value);
// 	}

// 	struct txml_node *current_entry = NULL;
// 	struct txml_node *current_sequence = NULL;
// 	struct txml_node *current_accession = NULL;
// 	struct txml_node *current_feature = NULL;
// 	struct txml_node *current_position = NULL;
// 	//struct txml_node *current_end = NULL;
	
// 	const char *ftype[] = { "feature", "@type" };
// 	while ((current_entry = txml_find(nodes, current_entry, TXML_ELEMENT, "entry", NULL, 1))) {
// 		printf("%s\n", current_entry->name);
// 		current_accession = txml_find(current_entry, current_accession, TXML_ELEMENT, "accession", NULL, 0);
// 		if (current_accession) {
// 			printf("%s\n", current_accession->value);
// 		}
// 		current_sequence = txml_find(current_entry, current_sequence, TXML_ELEMENT, "sequence", NULL, 0);
// 		if (current_sequence) {
// 			printf("%s\n", current_sequence->value);
// 		}

// 		while ((current_feature = txml_get(current_entry, current_feature, count(ftype), ftype))) {
// 			if (current_feature) {
// 				printf("%s", current_feature->value);
// 				//printf("%s\n", current_feature->parent->name);
// 				while ((current_position = txml_find(current_feature->parent, current_position, TXML_ATTRIBUTE, "position", NULL, 1))) {
// 					if (current_position) {
// 						printf("\t%s", current_position->value);
// 					}
// 				}
// 				printf("\n");
// 			}
// 		}
// 	}

// 	free(doc);
// 	free(nodes);
// 	printf("STATIC\n");

// 	printf("DYNAMIC\n");
// 	doc = txml_parse_file(argv[1], &nodes); /* The XML document as a zero-terminated string */

// 	current_entry = NULL;
// 	current_sequence = NULL;
// 	current_accession = NULL;
// 	current_feature = NULL;
// 	current_position = NULL;
// 	//struct txml_node *current_end = NULL;
	
// 	//const char *ftype[] = { "feature", "@type" };
// 	while ((current_entry = txml_find(nodes, current_entry, TXML_ELEMENT, "entry", NULL, 1))) {
// 		printf("%s\n", current_entry->name);
// 		current_accession = txml_find(current_entry, current_accession, TXML_ELEMENT, "accession", NULL, 0);
// 		if (current_accession) {
// 			printf("%s\n", current_accession->value);
// 		}
// 		current_sequence = txml_find(current_entry, current_sequence, TXML_ELEMENT, "sequence", NULL, 0);
// 		if (current_sequence) {
// 			printf("%s\n", current_sequence->value);
// 		}

// 		while ((current_feature = txml_get(current_entry, current_feature, count(ftype), ftype))) {
// 			if (current_feature) {
// 				printf("%s", current_feature->value);
// 				//printf("%s\n", current_feature->parent->name);
// 				while ((current_position = txml_find(current_feature->parent, current_position, TXML_ATTRIBUTE, "position", NULL, 1))) {
// 					if (current_position) {
// 						printf("\t%s", current_position->value);
// 					}
// 				}
// 				printf("\n");
// 			}
// 		}
// 	}

// 	printf("DYNAMIC\n");

// 	free(doc);
// 	free(nodes);

// 	return 0;
// }

#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include "util.h"
#include "miniz.h" // Include miniz for ZIP handling

#define TXML_DEFINE
#include "txml.h"  // Include txml for XML parsing

#define LEN(a)		sizeof(a) / sizeof(a[0]) 
#define TEMPFILE	"/tmp/doctxt-temp.txt" 

void
writetofile(char *out_file_path, char *data)
{
	FILE *out_file;
	out_file = fopen(out_file_path, "w");
	fprintf(out_file, "%s", data);
	fclose(out_file);
}

char *xml_unescape(const char *in) {
    char *out = malloc(strlen(in) * 2); // generous size
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

int
readzip(const char *path, char *filename)
{
	// Open the zip file using miniz
	mz_zip_archive zip;
	memset(&zip, 0, sizeof(zip));
	if (!mz_zip_reader_init_file(&zip, path, 0)) {
		die("Unable to open zip: %s", path);
	}

	// Find the file inside the zip archive
	int file_index = mz_zip_reader_locate_file(&zip, filename, NULL, 0);
	if (file_index < 0) {
		mz_zip_reader_end(&zip);
		return 0; // File not found
	}

	// Extract the file to memory
	size_t file_size;
	void *file_data = mz_zip_reader_extract_to_heap(&zip, file_index, &file_size, 0);
	if (!file_data) {
		mz_zip_reader_end(&zip);
		die("Failed to extract file from zip");
	}

	// Write the extracted content to the temporary file
	writetofile(TEMPFILE, (char *)file_data);

	// Clean up
	mz_free(file_data);
	mz_zip_reader_end(&zip);
	return 1; // Success
}

void extract_text_nodes(struct txml_node *parent, FILE *outfile)
{
    struct txml_node *node_t = NULL;
    while ((node_t = txml_find(parent, node_t, TXML_ELEMENT, "w:t", NULL, 1))) {
        char *text = node_t->value ? xml_unescape(node_t->value) : NULL;
        if (text) {
            fprintf(outfile, "%s", text);
            free(text);
        }
    }
}

void extract_table(struct txml_node *table, FILE *outfile)
{
    struct txml_node *row = NULL;
    
    // Traverse all table rows (w:tr)
    while ((row = txml_find(table, row, TXML_ELEMENT, "w:tr", NULL, 0))) {
        struct txml_node *cell = NULL;
        int first_cell = 1;
        
        // Traverse all table cells (w:tc) in the row
        while ((cell = txml_find(row, cell, TXML_ELEMENT, "w:tc", NULL, 0))) {
            if (!first_cell) {
                fprintf(outfile, "\t");
            }
            first_cell = 0;
            
            // Extract all paragraphs in the cell
            struct txml_node *para = NULL;
            int first_para = 1;
            int has_content = 0;
            while ((para = txml_find(cell, para, TXML_ELEMENT, "w:p", NULL, 0))) {
                if (!first_para) {
                    fprintf(outfile, " ");
                }
                first_para = 0;
                extract_text_nodes(para, outfile);
                has_content = 1;
            }
            // Empty cells are represented by nothing between tabs (handled by separator logic)
        }
        fprintf(outfile, "\n");
    }
}

void parsexml(char *path, FILE *outfile)
{
    struct txml_node *nodes = NULL;

    // Parse the XML data
    char *xml_data = txml_parse_file(path, &nodes);
    if (!xml_data) {
    	die("Error reading XML");
    }

    // Find the body node
    struct txml_node *node_body = NULL;
    node_body = txml_find(nodes, node_body, TXML_ELEMENT, "w:body", NULL, 1);
    if (!node_body) {
        die("No body element found in XML");
    }

    // Process paragraphs and tables
    // Note: Due to txml API limitations, we extract paragraphs first, then tables.
    // This may not preserve the exact document order if tables and paragraphs are interleaved.
    // For most documents this is acceptable as tables are typically grouped.
    struct txml_node *node_p = NULL, *node_tbl = NULL;
    
    // First, extract all paragraphs
    while ((node_p = txml_find(node_body, node_p, TXML_ELEMENT, "w:p", NULL, 0))) {
        extract_text_nodes(node_p, outfile);
        fprintf(outfile, "\n");
    }
    
    // Then extract all tables
    while ((node_tbl = txml_find(node_body, node_tbl, TXML_ELEMENT, "w:tbl", NULL, 0))) {
        extract_table(node_tbl, outfile);
    }

    // Clean up the allocated memory for nodes and XML data
    free(nodes);
    free(xml_data);
}

void parsecomments(char *path, FILE *outfile)
{
    struct txml_node *nodes = NULL;

    // Parse the XML data
    char *xml_data = txml_parse_file(path, &nodes);
    if (!xml_data) {
        // No comments file exists - silently return
        return;
    }

    // Find the comments root element
    struct txml_node *comments_root = NULL;
    comments_root = txml_find(nodes, NULL, TXML_ELEMENT, "w:comments", NULL, 0);
    if (!comments_root) {
        // No comments element found
        free(nodes);
        free(xml_data);
        return;
    }

    // Find all comment nodes within the comments element
    struct txml_node *comment = NULL;
    while ((comment = txml_find(comments_root, comment, TXML_ELEMENT, "w:comment", NULL, 0))) {
        // Get comment author
        struct txml_node *author_attr = txml_find(comment, NULL, TXML_ATTRIBUTE, "w:author", NULL, 0);
        if (author_attr && author_attr->value) {
            fprintf(outfile, "[%s]: ", author_attr->value);
        } else {
            fprintf(outfile, "[Unknown]: ");
        }
        
        // Extract text from all paragraphs in the comment
        struct txml_node *para = NULL;
        while ((para = txml_find(comment, para, TXML_ELEMENT, "w:p", NULL, 0))) {
            extract_text_nodes(para, outfile);
        }
        fprintf(outfile, "\n");
    }

    // Clean up the allocated memory for nodes and XML data
    free(nodes);
    free(xml_data);
}

void
usage()
{
	die("usage: doctxt infile [-o outfile] [-c]");
}

int
main(int argc, char *argv[])
{
	FILE *outfile = NULL;
	char *outfilename = "out.txt";
	char *infilename = NULL;
	int comments_only = 0;

	if (argc < 2) {
		usage();
	}

	// Parse all arguments
	for (int i = 1; i < argc; i++) {
		if (!strcmp(argv[i], "-v")) {
			puts("doctxt-"VERSION);
			return 0;
		} else if (!strcmp(argv[i], "-c")) {
			comments_only = 1;
		} else if (!strcmp(argv[i], "-o")) {
			if (i + 1 >= argc) {
				usage();
			}
			outfilename = argv[i + 1];
			i++;
		} else if (argv[i][0] == '-') {
			// Unknown flag
			usage();
		} else {
			// Input filename
			if (infilename != NULL) {
				usage(); // Multiple input files not allowed
			}
			infilename = argv[i];
		}
	}

	// Check if input filename was provided
	if (infilename == NULL) {
		usage();
	}

	// Open output file for writing
	outfile = fopen(outfilename, "wt");
	if (!outfile) {
		die("Unable to open output file: %s", outfilename);
	}

	if (comments_only) {
		// Extract only comments
		if (readzip(infilename, "word/comments.xml")) {
			parsecomments(TEMPFILE, outfile);
			// Remove the temporary file
			if (remove(TEMPFILE) != 0) {
				die("Unable to delete tempfile");
			}
		}
		// If no comments file, output file will be empty
	} else {
		// Extract document content (text and tables)
		if (!readzip(infilename, "word/document.xml")) {
			fclose(outfile);
			die("File not found in zip: word/document.xml");
		}
		parsexml(TEMPFILE, outfile);
		// Remove the temporary file
		if (remove(TEMPFILE) != 0) {
			die("Unable to delete tempfile");
		}
	}

	// Close the output file
	fclose(outfile);

	return 0;
}
