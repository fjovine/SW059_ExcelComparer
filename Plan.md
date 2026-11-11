# SW059_ExcelComparer

This file loads two text files and creates an excel sheet containing each line on facing cells.

If for instance FileA.txt contains the following lines
A1
A2
A3

and FileB.txt contains the following lines
B1
B2
B3
B4

the excel output file will contain the following table

| # |FileA|FileB|
|:- |:----|:----|
|1|A1|B1|
|2|A2|B2|
|3|A3|B3|
| |  |B4|

So the first column contains the line numbers of the first file, 
the second column contains the corresponding line of the first file while
the third column contains the lines of the second file.

It may be useful to align the content of a log file with the source code of the 
program executed

## Synopsis

dotnet run <filea> <fileb> [--encoding=windows-1252]

So the first parameter is the filename of file A, the second is the filename of fileB and an optional parameter permits to change the encoding of both files

The outputfile is <filea>_<fileb>.xlsx

In this case filea and fileb are the filenames of the corresponding files without extension