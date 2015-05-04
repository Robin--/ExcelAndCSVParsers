# ExcelAndCSVParsers

This is a simple C# library that cleans and parses CSV and Excel files. 

The DLL reads both Excel and CSV files, cleans, standarsises them, and outputs a string based class containing the CSV data.
The class can also save Excel and CSV files.

The CSV parser also has the option to remove all UNIX characters.

The DLL was written as a DLL plugin for Warewolf ESB (https://github.com/Warewolf-ESB/Warewolf-ESB) to clean up input CSV to make it simple to parse in Warewolf.
The DLL can also read and save Excel files into CSV data bojects, for parsing in Warewolf.

# The CSV Class
The Public method, Parse, takes in 4 values:

string inputFilename - The full UNC path to the CSV input file. e.g. C:\data\inputFile.csv .

string splitCharacter - the field split character. e.g. , .

string quoteCharacter - the text field quote character to use e.g. " .

bool trimWhiteSpaces - remove leading and trailing whitespaces from all fields e.g. True or False.

bool useFields - Include the field header row in the output  e.g. True or False.
