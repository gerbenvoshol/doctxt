/*
 * xlsx2md - convert Excel 2007 files to Markdown tables
 * 
 * Based on xlsx2csv functionality
 * Modified to output Markdown format
 * 
 * Copyright (C) 2024 Gerben Voshol
 * 
 * This program is free software; you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation; either version 2 of the License, or
 * (at your option) any later version.
 */

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include "util.h"
#include "miniz.h"

#define TXML_DEFINE
#include "txml.h"

#define VERSION "0.1"
#define BUFFSIZE 40960
#define MAX_NODES (1024 * 1024)
#define MAX_COLS 256

typedef struct {
	FILE *outf;
	char **shrdstr_array;
	int shrdstr_cnt;
	int sheet_num_rows, sheet_num_cols;
	int current_row, current_col;
	char **row_data;  /* Store current row data for formatting */
	int *col_widths;  /* Track column widths for alignment */
	int is_first_row; /* Track if we're processing the first row (header) */
} XLSXCtx;

static void
output_markdown_cell(FILE *out, const char *z)
{
	if (z == NULL || *z == '\0') {
		/* empty cell */
		return;
	}
	
	/* Escape pipe characters in Markdown tables */
	for (const char *p = z; *p; p++) {
		if (*p == '|') {
			fprintf(out, "\\|");
		} else if (*p == '\n' || *p == '\r') {
			/* Replace newlines with spaces in table cells */
			fputc(' ', out);
		} else {
			fputc(*p, out);
		}
	}
}

static void
excelcolrow(char *string, int *outcol, int *outrow)
{
	int i, col;

	col = 0;
	for (i = 0; i < strlen(string); i++) {
		if (isalpha(string[i])) {
			col = col * 26 + ((toupper(string[i])) - 'A' + 1);
		} else
			break;
	}
	*outcol = col;
	*outrow = atoi(string + i);
}

static void
rangecolrow(char *string, int *outcol, int *outrow)
{
	char *coloninstr = strchr(string, ':');
	if (coloninstr) {
		string = coloninstr + 1;
	}
	excelcolrow(string, outcol, outrow);
}

static void
parse_shared_strings(XLSXCtx *ctx, char *xml_data)
{
	struct txml_node *nodes = xmalloc(sizeof(struct txml_node) * MAX_NODES);
	char *tail = txml_parse(xml_data, MAX_NODES, nodes);
	
	if (tail) {
		fprintf(stderr, "Warning: XML parsing incomplete for sharedStrings.xml\n");
	}

	/* Find the sst element to get uniqueCount */
	struct txml_node *sst = txml_find(nodes, NULL, TXML_ELEMENT, "sst", NULL, 1);
	if (sst) {
		struct txml_node *attr = txml_find(sst, NULL, TXML_ATTRIBUTE, "uniqueCount", NULL, 0);
		if (attr && attr->value) {
			ctx->shrdstr_cnt = atoi(attr->value);
			ctx->shrdstr_array = ecalloc(ctx->shrdstr_cnt, sizeof(char *));
		}
	}

	/* Parse all shared strings */
	int str_idx = 0;
	struct txml_node *si = NULL;
	while ((si = txml_find(nodes, si, TXML_ELEMENT, "si", NULL, 1))) {
		char buffer[BUFFSIZE] = {0};
		
		/* Find <t> text nodes within this <si> and concatenate them */
		struct txml_node *t = NULL;
		while ((t = txml_find(si, t, TXML_ELEMENT, "t", NULL, 1))) {
			struct txml_node *text = txml_find(t, NULL, TXML_TEXT, NULL, NULL, 0);
			if (text && text->value) {
				strncat(buffer, text->value, BUFFSIZE - strlen(buffer) - 1);
			}
		}
		
		if (str_idx < ctx->shrdstr_cnt) {
			ctx->shrdstr_array[str_idx++] = strdup(buffer);
		}
	}

	free(nodes);
}

static void
output_row_as_markdown(XLSXCtx *ctx)
{
	/* Output row with Markdown table format */
	fprintf(ctx->outf, "|");
	for (int i = 0; i < ctx->sheet_num_cols; i++) {
		fprintf(ctx->outf, " ");
		if (ctx->row_data[i]) {
			output_markdown_cell(ctx->outf, ctx->row_data[i]);
		}
		fprintf(ctx->outf, " |");
	}
	fprintf(ctx->outf, "\n");
	
	/* If this was the first row (header), output separator */
	if (ctx->is_first_row) {
		fprintf(ctx->outf, "|");
		for (int i = 0; i < ctx->sheet_num_cols; i++) {
			fprintf(ctx->outf, " --- |");
		}
		fprintf(ctx->outf, "\n");
		ctx->is_first_row = 0;
	}
	
	/* Clear row data */
	for (int i = 0; i < ctx->sheet_num_cols; i++) {
		if (ctx->row_data[i]) {
			free(ctx->row_data[i]);
			ctx->row_data[i] = NULL;
		}
	}
}

static void
parse_sheet(XLSXCtx *ctx, char *xml_data)
{
	struct txml_node *nodes = xmalloc(sizeof(struct txml_node) * MAX_NODES);
	char *tail = txml_parse(xml_data, MAX_NODES, nodes);
	
	if (tail) {
		fprintf(stderr, "Warning: XML parsing incomplete for worksheet\n");
	}

	/* Find dimension to get sheet size */
	struct txml_node *dimension = txml_find(nodes, NULL, TXML_ELEMENT, "dimension", NULL, 1);
	if (dimension) {
		struct txml_node *ref_attr = txml_find(dimension, NULL, TXML_ATTRIBUTE, "ref", NULL, 0);
		if (ref_attr && ref_attr->value) {
			rangecolrow((char *)ref_attr->value, &ctx->sheet_num_cols, &ctx->sheet_num_rows);
		}
	}

	/* Allocate row data storage */
	if (ctx->sheet_num_cols > MAX_COLS) {
		ctx->sheet_num_cols = MAX_COLS;
	}
	ctx->row_data = ecalloc(ctx->sheet_num_cols, sizeof(char *));
	ctx->is_first_row = 1;

	/* Process rows */
	struct txml_node *row = NULL;
	while ((row = txml_find(nodes, row, TXML_ELEMENT, "row", NULL, 1))) {
		/* Process cells in this row */
		struct txml_node *cell = NULL;
		while ((cell = txml_find(row, cell, TXML_ELEMENT, "c", NULL, 0))) {
			int lookup_v = 0;
			
			/* Get cell attributes */
			struct txml_node *r_attr = txml_find(cell, NULL, TXML_ATTRIBUTE, "r", NULL, 0);
			struct txml_node *t_attr = txml_find(cell, NULL, TXML_ATTRIBUTE, "t", NULL, 0);
			
			if (t_attr && t_attr->value && *t_attr->value == 's') {
				lookup_v = 1;
			}
			
			if (r_attr && r_attr->value) {
				int col, row_num;
				excelcolrow((char *)r_attr->value, &col, &row_num);
				
				/* Get cell value */
				struct txml_node *v = txml_find(cell, NULL, TXML_ELEMENT, "v", NULL, 0);
				if (v) {
					struct txml_node *text = txml_find(v, NULL, TXML_TEXT, NULL, NULL, 0);
					if (text && text->value && col > 0 && col <= ctx->sheet_num_cols) {
						if (lookup_v && ctx->shrdstr_array) {
							int idx = atoi(text->value);
							if (idx >= 0 && idx < ctx->shrdstr_cnt) {
								ctx->row_data[col - 1] = strdup(ctx->shrdstr_array[idx]);
							}
						} else {
							ctx->row_data[col - 1] = strdup(text->value);
						}
					}
				}
			}
		}
		
		/* Output the complete row */
		output_row_as_markdown(ctx);
	}

	free(ctx->row_data);
	free(nodes);
}

static void
usage(void)
{
	fprintf(stderr, "usage: xlsx2md [-if input.xlsx] [-sh sheet_id] [-of output.md] [-v] [-h]\n");
	fprintf(stderr, "  -if FILE      input spreadsheet in Excel 2007 format\n");
	fprintf(stderr, "  -sh NUMBER    sheet number (default: 1)\n");
	fprintf(stderr, "  -of FILE      output Markdown file (default: stdout)\n");
	fprintf(stderr, "  -v            display version information\n");
	fprintf(stderr, "  -h            display this help message\n");
	exit(1);
}

int
main(int argc, char *argv[])
{
	XLSXCtx ctx = {0};
	char *input_file = NULL;
	char *output_file = NULL;
	int sheet_num = 1;
	size_t size;
	void *data;
	char sheetname[64];

	/* Parse arguments */
	for (int i = 1; i < argc; i++) {
		if (strcmp(argv[i], "-if") == 0 && i + 1 < argc) {
			input_file = argv[++i];
		} else if (strcmp(argv[i], "-sh") == 0 && i + 1 < argc) {
			sheet_num = atoi(argv[++i]);
			if (sheet_num < 1) sheet_num = 1;
		} else if (strcmp(argv[i], "-of") == 0 && i + 1 < argc) {
			output_file = argv[++i];
		} else if (strcmp(argv[i], "-v") == 0) {
			printf("xlsx2md version %s\n", VERSION);
			return 0;
		} else if (strcmp(argv[i], "-h") == 0) {
			usage();
		} else {
			fprintf(stderr, "Unknown option: %s\n", argv[i]);
			usage();
		}
	}

	if (!input_file) {
		fprintf(stderr, "Error: Missing input file (-if)\n");
		usage();
	}

	/* Open output file */
	if (output_file) {
		ctx.outf = fopen(output_file, "w");
		if (!ctx.outf) {
			die("Could not open output file: %s\n", output_file);
		}
	} else {
		ctx.outf = stdout;
	}

	/* Parse shared strings */
	data = mz_zip_extract_archive_file_to_heap(input_file, "xl/sharedStrings.xml", &size, 0);
	if (data) {
		parse_shared_strings(&ctx, (char *)data);
		mz_free(data);
	}

	/* Parse sheet */
	snprintf(sheetname, sizeof(sheetname), "xl/worksheets/sheet%d.xml", sheet_num);
	data = mz_zip_extract_archive_file_to_heap(input_file, sheetname, &size, 0);
	if (!data) {
		die("Could not read sheet %d\n", sheet_num);
	}
	
	parse_sheet(&ctx, (char *)data);
	mz_free(data);

	/* Cleanup */
	if (ctx.shrdstr_array) {
		for (int i = 0; i < ctx.shrdstr_cnt; i++) {
			free(ctx.shrdstr_array[i]);
		}
		free(ctx.shrdstr_array);
	}

	if (output_file) {
		fclose(ctx.outf);
	}

	return 0;
}
