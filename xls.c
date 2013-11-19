#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <unistd.h>

#include <xls.h>

static char  *stringSeparator = "";
static char *lineSeparator = "\n";
static char *fieldSeparator = ",";
static char *encoding = "UTF-8";
char *lines;
static double size = 4096;


static void OutputString(const char *string);
static void OutputNumber(const double number);


static void appendString(char * str)
{
	if (strlen(lines) > size){
		lines = realloc(lines,2 * size);
		size = 2 *size;
	} else {
		strcat(lines,str);
	}
}

int readSheet(char * filename, char * sheetName) {
	xlsWorkBook* pWB;
	xlsWorkSheet* pWS;
	unsigned int i;
    int justList = 0;
    lines = (char *)malloc(4096);
    memset(lines,0x00,4096);

	//fprintf(stderr, "DIR: %s\n\n", getcwd(NULL, 1024));
	//printf("%s\n", "readSheet start");
	struct st_row_data* row;
	WORD cellRow, cellCol;

	// open workbook, choose standard conversion
	pWB = xls_open(filename, encoding);
	if (!pWB) {
		fprintf(stderr, "File not found");
		fprintf(stderr, "\n");
		return EXIT_FAILURE;
	}

	// check if the requested sheet (if any) exists
	if (sheetName[0]) {
		for (i = 0; i < pWB->sheets.count; i++) {
			if (strcmp(sheetName, (char *)pWB->sheets.sheet[i].name) == 0) {
				break;
			}
		}

		if (i == pWB->sheets.count) {
			fprintf(stderr, "Sheet \"%s\" not found", sheetName);
			fprintf(stderr, "\n");
			return EXIT_FAILURE;
		}
	}

	// process all sheets
	for (i = 0; i < pWB->sheets.count; i++) {
		int isFirstLine = 1;

		// check if this the sheet we want
		if (sheetName[0]) {
			if (strcmp(sheetName, (char *)pWB->sheets.sheet[i].name) != 0) {
				continue;
			}
		}

		// open and parse the sheet
		pWS = xls_getWorkSheet(pWB, i);
		xls_parseWorkSheet(pWS);

		// process all rows of the sheet
		for (cellRow = 0; cellRow <= pWS->rows.lastrow; cellRow++) {
			int isFirstCol = 1;
			row = xls_row(pWS, cellRow);

			// process cells
			if (!isFirstLine) {
				appendString(lineSeparator);
				// printf("%s", lineSeparator);
			} else {
				isFirstLine = 0;
			}

			for (cellCol = 0; cellCol <= pWS->rows.lastcol; cellCol++) {
                //printf("Processing row=%d col=%d\n", cellRow+1, cellCol+1);

				xlsCell *cell = xls_cell(pWS, cellRow, cellCol);

				if ((!cell) || (cell->isHidden)) {
					continue;
				}

				if (!isFirstCol) {
					appendString(fieldSeparator);
					//printf("%s", fieldSeparator);
				} else {
					isFirstCol = 0;
				}

				// display the colspan as only one cell, but reject rowspans (they can't be converted to CSV)
				if (cell->rowspan > 1) {
					fprintf(stderr, "Warning: %d rows spanned at col=%d row=%d: output will not match the Excel file.\n", cell->rowspan, cellCol+1, cellRow+1);
				}

				// display the value of the cell (either numeric or string)
				if (cell->id == 0x27e || cell->id == 0x0BD || cell->id == 0x203) {
					OutputNumber(cell->d);
				} else if (cell->id == 0x06) {
                    // formula
					if (cell->l == 0) // its a number
					{
						OutputNumber(cell->d);
					} else {
						if (!strcmp((char *)cell->str, "bool")) // its boolean, and test cell->d
						{
							OutputString((int) cell->d ? "true" : "false");
						} else if (!strcmp((char *)cell->str, "error")) // formula is in error
						{
							OutputString("*error*");
						} else // ... cell->str is valid as the result of a string formula.
						{
							OutputString((char *)cell->str);
						}
					}
				} else if (cell->str != NULL) {
					OutputString((char *)cell->str);
				} else {
					OutputString("");
				}
			}
		}
		xls_close_WS(pWS);
	}

	xls_close(pWB);
	//printf("%s\n", lines);
	return EXIT_SUCCESS;
}

// Output a CSV String (between double quotes)
// Escapes (doubles)" and \ characters
static void OutputString(const char *string) {
	const char *str;
	char tmp[20];
	sprintf(tmp,"%s",stringSeparator);
	appendString(tmp);
	//printf("%c", stringSeparator);
	for (str = string; *str; str++) {
		if (*str == '\\') {
			appendString("\\\\");
			//printf("\\\\");
		} else {
			sprintf(tmp,"%c", *str);
			appendString(tmp);
			//printf("%c", *str);
		}
	}
	sprintf(tmp,"%s",stringSeparator);
	appendString(tmp);
	//printf("%c", stringSeparator);
}

// Output a CSV Number
static void OutputNumber(const double number) {
	char tmp[20];
	sprintf(tmp,"%.15g", number);
	appendString(tmp);
	//printf("%.15g", number);
}
