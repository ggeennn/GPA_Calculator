/*
 * Course:		IPC-NDD
 * Name:	 	Yicheng Wang
 * Student ID:	150868206
 * Problems Solved: Topsilver(A2,easy), Top10Medal(C3,hard)
 */
 
/*Structure representing Olympic medal data*/
struct MedalCount {
	char country[50];
	int year;
	char season[10]; // "Summer" or "Winter"
	int gold;
	int silver;
	int bronze;
	int total;
};

/*please check the function DO states before each function's implementation*/
void UsrInput(int* year, char* season);

void ReadMedalFile(const char* filename, struct MedalCount* data, int* size, int year, const char* season);

void Topsilver(struct MedalCount* data, int size, int* maxsilver, char* maxcountry);

void Top10Medal(struct MedalCount* data, int size, struct MedalCount* top10, int* count);

void PrintSummary(int year, const char* season, struct MedalCount* data, int size);
