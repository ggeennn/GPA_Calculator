/*
 * Course:		IPC-NDD
 * Name:	 	Yicheng Wang
 * Student ID:	150868206
 * Problems Solved: Topsilver(A2,easy), Top10Medal(C3,hard)
 */
 
#define _CRT_SECURE_NO_WARNINGS
#include <stdio.h>
#include <string.h>		
#include "project.h"
/*------------------------------------------------------------
 * This function prompts the user for a year and season input.
 *
 * Parameters:
 *   int* year: Pointer to an integer where the input year will be stored.
 *   char* season: Pointer to a string where the input season will be stored.
 *
 * Limitations:
 *   - The year must be a non-negative integer.
 *   - The season must be either "Summer" or "Winter".
 *
 * Return Value: None (void).
 */
void UsrInput(int* year, char* season) {
	do {
		printf("Please enter the year: ");
		if (scanf(" %d", year) != 1) {
			while (getchar() != '\n'); // Clear invalid input
			printf("Invalid input, please try again.\n");
		}
	} while (*year < 0); // Ensure year is non-negative
	
	int i = 0;
	do {
		printf("Type in: 1(summer) / 2(winter) ? ");
		if (scanf(" %d", &i)) { // Validate season input
			if (i == 1) {
				strcpy(season, "Summer");
			}
			else if (i == 2) {
				strcpy(season, "Winter");
			}
			else {
				printf("Invalid input, please enter 1 or 2.\n");
			}
		}
		else {
			while (getchar() != '\n');
			printf("Invalid input, please enter 1 or 2.\n");
		}
	} while (i != 1 && i != 2); // Ensure valid season input
}

/*-----------------------------------------------------------------------------------------
 * This function reads the medal data from a CSV file and populates the provided data array.
 *
 * Parameters:
 *   const char* filename: The name of the file to read from.
 *   struct MedalCount* data: Pointer to an array of MedalCount structures where the data will be stored.
 *   int* size: Pointer to an integer where the number of records read will be stored.
 *   int year: The year to filter the records.
 *   const char* season: The season to filter the records.
 *
 * Limitations:
 *   - The function assumes the file exists and is formatted correctly.
 *   - The data array must be large enough to hold the expected number of records.
 *
 * Return Value: None (void).
 */ 
void ReadMedalFile(const char* filename, struct MedalCount* data, int* size, int year, const char* season) {
	FILE* fp = NULL;
	char line[300];  // Buffer for reading lines

	fp = fopen(filename, "r");
	if (fp != NULL) {
		fgets(line, sizeof(line), fp); // Skip header line

		int i = 0; // Counter for data stored
		int c = 1; // Line counter
		while (fgets(line, sizeof(line), fp)) {
			c++;
			
			/*Checking if the line of file excceeds the buffer size*/
			if (strchr(line, '\n') == NULL && strlen(line) == sizeof(line) - 1) {
				printf("Incomplete line NO.%d detected (truncated)...\nStatistics below might be incorrect:\n\n", c);
				int ch;
				do {
					ch = fgetc(fp);
				} while (ch != '\n' && ch != EOF); // Skip to the end of the line
			}

			struct MedalCount temp = { "NULL",0,"NULL",0,0,0,0 }; // For storing current read line
			char noc[30]; // For handling specific data cases (1956 Equestrian)

			sscanf(line, "%*[^,],%*d,%d,%[^,],%[^,],%d,%d,%d,%d", &temp.year, temp.country, noc, &temp.gold, &temp.silver, &temp.bronze, &temp.total);

			if (temp.year == year && strstr(line, season) != NULL) {

				/*Adjust medal counts for specific cases(1956 Equestrian)*/
				if (year == 1956 && strcmp(season, "Summer") == 0) {
					if (strcmp(noc, "SWE") == 0) {
						temp.gold += 3;
					}
					if (strcmp(noc, "GER") == 0) {
						temp.gold += 2;
						temp.silver += 3;
						temp.bronze += 1;
					}
					if (strcmp(noc, "GBR") == 0) {
						temp.gold += 1;
						temp.bronze += 2;
					}
					if (strcmp(noc, "ITA") == 0) {
						temp.silver += 2;
						temp.bronze += 1;
					}
					if (strcmp(noc, "DEN") == 0) {
						temp.silver += 1;
					}
					if (strcmp(noc, "CAN") == 0 || strcmp(noc, "CAN") == 0) {
						temp.bronze += 1;
					}
				}

				 /* Store the valid medal count in the data array */
                if (i < 150) { // Ensure we do not exceed the array bounds
                    strcpy(data[i].country, temp.country);
                    data[i].year = temp.year;
                    strcpy(data[i].season, season);
                    data[i].gold = temp.gold;
                    data[i].silver = temp.silver;
                    data[i].bronze = temp.bronze;
                    data[i].total = temp.gold + temp.silver + temp.bronze;
                    i++;
                } else {
                    printf("Warning: Data array size is insufficient. Consider increasing the size.\n");
                }
			}
		}
		*size = i; // Update the size of the data
		fclose(fp);
	}
	else {
		printf("Error: Could not open file.\n");
	}
}

/*----------------------------------------------------------------------------
 * This function finds the country with the maximum number of silver medals.
 *
 * Parameters:
 *   struct MedalCount* data: Pointer to an array of MedalCount structures containing medal data.
 *   int size: The number of records in the data array.
 *   int* maxsilver: Pointer to an integer where the maximum silver count will be stored.
 *   char* maxcountry: Pointer to a character array where the country with the most silver medals will be stored.
 *
 * Limitations:
 *   - The data array must contain valid MedalCount records.
 *
 * Return Value: None (void).
 */
void Topsilver(struct MedalCount* data, int size, int* maxsilver, char* maxcountry) {
	int i;
	for (i = 0; i < size; i++) {
		if (data[i].silver > *maxsilver) { // Find the country with the most silver medals
			*maxsilver = data[i].silver;
			strcpy(maxcountry, data[i].country);
		}
		else if (data[i].silver == *maxsilver) { // Append country name if tied
			/*Check for potential overflow when concatenating country names*/
			if (strlen(maxcountry) + strlen(data[i].country) + 1 < sizeof(maxcountry)) {
				strcat(maxcountry, " ");  
				strcat(maxcountry, data[i].country); 
			}
			else {
				printf("Warning: Country name overflow detected when adding: %s\n", data[i].country);
			}
		}
	}
}

/*---------------------------------------------------------------------
 * This function determines the top 10 countries based on total medals.
 *
 * Parameters:
 *   struct MedalCount* data: Pointer to an array of MedalCount structures containing medal data.
 *   int size: The number of records in the data array.
 *   struct MedalCount* top10: Pointer to an array of MedalCount structures where the top 10 countries will be stored.
 *   int* count: Pointer to an integer where the number of top countries will be stored.
 *
 * Limitations:
 *   - The data array must contain valid MedalCount records.
 *
 * Return Value: None (void).
 */
void Top10Medal(struct MedalCount* data, int size, struct MedalCount* top10, int* count) {
	struct MedalCount temp; // Temporary structure for swapping
	
	/*Swap data of top 10 countries to the beginning of data array*/
	int i;
	for (i = 0; i < (size > 10 ? 10 : size); i++) {
		int j;
		for (j = i + 1; j < size; j++) {

			if ( // Sorting rule based on total medals, then gold, silver, and bronze
				data[j].total > data[i].total ||
				(data[j].total == data[i].total && data[j].gold > data[i].gold) ||
				(data[j].total == data[i].total && data[j].gold == data[i].gold && data[j].silver > data[i].silver) ||
				(data[j].total == data[i].total && data[j].gold == data[i].gold && data[j].silver == data[i].silver && data[j].bronze > data[i].bronze)
				) {
				
				/*Swap data[i] and data[j]*/
				temp = data[i];
				data[i] = data[j];
				data[j] = temp;
			}
		}
	}

	*count = size < 10 ? size : 10; // Set count to the number of top countries
	for (i = 0; i < *count; i++) {
		top10[i] = data[i]; // Copy top 10 countries to the top10 array
	}
}

/*------------------------------------------------------------------------------------
 * This function prints a summary of the medal records for a specific year and season.
 *
 * Parameters:
 *   int year: The year for which the summary is printed.
 *   const char* season: The season for which the summary is printed.
 *   struct MedalCount* data: Pointer to an array of MedalCount structures containing medal data.
 *   int size: The number of records in the data array.
 *
 * Limitations:
 *   - The data array must contain valid MedalCount records.
 *
 * Return Value: None (void).
 */
void PrintSummary(int year, const char* season, struct MedalCount* data, int size) {
	int maxSilver = 0;
	char maxCountry[100] = { 0 }; // Store the country names with most silver medals
	Topsilver(data, size, &maxSilver, maxCountry); // Get top silver winners

	printf("\n========================================\n");
	printf("    Olympic Medal Records Summary\n\n");
	printf("Total records: %d\n", size);
	printf("========================================\n");
	printf("Country(ies) with most silver medals:\n%s\n\n", maxCountry);
	printf("%-13s : %d\n", "Year", year);
	printf("%-13s : %s\n", "Season", season);
	printf("%-13s : %d\n", "Silver Medals", maxSilver);
	printf("========================================\n");

	struct MedalCount top10[10]; // Store top 10 winners
	int count = 0; // Store the number of top countries (may less than 10)
	Top10Medal(data, size, top10, &count); // Get top 10 medal winners

	printf("Top %d Countries Medal Histogram:\n\n", count);

	/*Find the longest country name to line up name of countires in histogram*/
	int maxNameLength = 0;
	int i;
	for (i = 0; i < count; i++) {
		if (strlen(top10[i].country) > maxNameLength) {
			maxNameLength = strlen(top10[i].country);
		}
	}
	/*Find the highest total medals to determine the scale of histogram*/
	int maxTotal = top10[0].total;
	for (i = 0; i < count; i++) {
		if (top10[i].total > maxTotal) {
			maxTotal = top10[i].total;
		}
	}
	int maxBlocks = 50;
	int scale = (maxTotal + maxBlocks - 1) / maxBlocks; // Scale formula

	/*Calculate and print blocks for histogram*/
	for (i = 0; i < count; i++) {
		printf("%*s: ", maxNameLength, top10[i].country); 
		int blocks = top10[i].total / scale; 
		int j;
		for (j = 0; j < blocks; j++) {
			printf("█"); 
		}
		printf(" (%d)\n\n", top10[i].total); 
	}
	printf("========================================\n");
}
