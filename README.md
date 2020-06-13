## PDF Analysis with UiPath

I have a lot of PDF documentation to read.  This automation utilises the native capabilities within UiPath to create an analysis of the words used within a PDF file.  Mapping the frequency that a word is used, whilst stripping out common english words.

### Steps

#### Set Variables

1) Requests the folder containing the PDFs for analysis
2) Requests the path to the downloaded wordCloud.txt (this contains the VBA code)

#### Analysis

##### For Each file in the provided directory;

3) Read PDF Text to string, assign to a string variable.
4) Generate single column DataTable from PDF Text, utilising 'space' as both a Column and NewLine Separator.  Assign to DataTable variable.
5) Assign DataTable column name to a string variable.
6) Count the Rows in the DataTable and assign that to a string variable
7) Create Excel file from PDF name with suffix '-report'
8) Right DataTable to Excel file starting from cell 'A1' on sheet named 'StratWords'
9) Creates a table in the sheet stratwords, using the variable from count rows action to determine the correct table size
10) Invokes VBA stored in wordCloud.txt at 'Punc' entry point to strip punctuation from the word list
11) Invoke VBA stored in wordCloud.txt at 'commonWords' entry point to remove common words from the word list
12) Inserts column 'Count'
13) Writes formula 'COUNTIF(A:A,A2)' to cells in Count column - this counts the frequency of a word
14) Inserts column 'Rank'
15) Writes formula 'RANK.EQ(B2,B:B,0)' to cells in the Rank column - this ranks the words from most to least frequently used.
16) Creates Pivot table call 'StratPiv' in new sheet called 'PDFPivot'
17) Invoke VBA stored in wordCloud.txt at 'PivotConfig' entry point to reformat the Pivot table
18) Saves the file

### Output Example

![Table Output Example](https://raw.githubusercontent.com/sconyard/UiPath_PDFwordAnalysis/master/images/PDFwordAnalysis_TableOutput.png)

![Pivot Output Example 1](https://raw.githubusercontent.com/sconyard/UiPath_PDFwordAnalysis/master/images/PDFwordAnalysis_PivotOutput1.png)

![Pivot Output Example 2](https://raw.githubusercontent.com/sconyard/UiPath_PDFwordAnalysis/master/images/PDFwordAnalysis_PivotOutput2.png)

### Support

No support offered or liability accepted, use this at your own risk.

This script was built and tested using;

Created using UiPath Studio Pro Community Edition 2020.4.1

No additional packages required for workflow
