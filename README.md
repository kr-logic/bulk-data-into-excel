This is my very first VBA project from my early automation days. It demonstrates 
fundamental logic and problem-solving skills, though modern implementations 
would utilize different performance techniques (see below).


DESCRIPTION
-----------
This VBA utility was developed to solve a critical data migration challenge: 
consolidating fragmented legacy data into a unified Master Dataset for 
financial analysis.

It functions as a lightweight ETL (Extract, Transform, Load) engine within 
Excel, capable of processing multitudes of semicolon separated data files
autonomously.

This version currently looks for files ending in .DAT format, but it can be
extended/changed easily to suit other needs (eg. .CSV).


KEY METRICS
-----------
* VOLUME: Successfully processed hundreds of source files in a single batch run.
* DATASET: Generated a consolidated master sheet of 70 000+ records.
* IMPACT: Replaced weeks of manual data entry with a single automated process.


THE ALGORITHM
-------------
1. DIRECTORY SCANNING
   Iterates through a user-selected folder to identify valid source files (.DAT).

2. TEXT PARSING
   Reads raw text streams line-by-line and splits them based on delimiters (;).

3. DATA TYPE ENFORCEMENT (CRITICAL)
   Standard Excel import often corrupts financial IDs (e.g., removing leading 
   zeros: "00123" -> "123"). This script strictly enforces Text formatting 
   for ID columns (Col 2, 9) while converting Amount columns (Col 6, 8) 
   to numeric values for calculation.

4. AGGREGATION
   Appends the cleaned data row-by-row into the Master Worksheet.


LEGACY NOTE & SELF-REFLECTION
-----------------------------
This tool was built when I first started automating financial processes. 
It utilizes direct cell writing (ws.Cells) inside the loop, which is reliable 
and easy to debug, but computationally expensive.

If I were to refactor this in the future for a production environment,
I would read the entire file into memory (array) and write to the sheet
in one operation to boost performance.

However, the logic presented here remains sound and successfully delivered 
the required business results at the time.


REQUIREMENTS
------------
* Microsoft Excel (VBA enabled)
* Source files in .DAT format (semicolon delimited)


USAGE
-----
1. Open the Excel file.
2. Run the 'ImportAndConsolidateFiles' subroutine.
3. A dialog window will appear. Select the folder containing your .DAT files.
4. The script will create a "Raw Data(VBA)" sheet and populate it.
5. A message box confirms when the process is complete.

---------------
AUTHOR: Princzinger Krisztián

Copyright (c) 2026 Princzinger Krisztián. All rights reserved.
