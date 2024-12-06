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

int main(void) {
	char choice; // store User's decision: continue or exit the program
	do {
		int year = -1;
		char season[10];
		UsrInput(&year, season); // Get user input for year and season

		char file[] = "Olympic_Games_Medal_Tally.csv";
		struct MedalCount data[150]; // Stores filtered data
		int size = 0; // Number of valid records
		ReadMedalFile(file, data, &size, year, season); // Read and filter medal data

		if (size > 0) {
			PrintSummary(year, season, data, size); // Print summary if data is available
		}
		else {
			printf("========================================\n");
			printf("    Olympic Medal Records Summary\n");
			printf("========================================\n");
			printf("No records found for the specified year and season.\n\n");
		}

		/*Repeat the program if user chooses 'y'*/
		do{
			printf("Keep searching? y(yes) / n(no) : ");
			scanf(" %c", &choice);
			while (getchar() != '\n'); // Clear input buffer
		} while (choice != 'y' && choice != 'Y' && choice != 'n' && choice != 'N');
		printf("========================================\n");
	} while (choice == 'y' || choice == 'Y');

	printf("Exiting...\n");
	return 0;
}