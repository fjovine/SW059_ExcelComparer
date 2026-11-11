# SW059_ExcelComparer

A small command-line tool that puts two text files side by side in an Excel sheet, line by line.

A typical use is aligning a log file with the source code of the program that produced it, so you can read the two together and add notes in Excel.

## Requirements

- [.NET 8 SDK](https://dotnet.microsoft.com/download) (or newer SDK with the .NET 8 runtime installed)

The only dependency, [ClosedXML](https://github.com/ClosedXML/ClosedXML), is restored automatically from NuGet on first build.

## Usage

```
dotnet run -- <filea> <fileb> [--encoding=<name>]
```

| Argument | Description |
|:--|:--|
| `<filea>` | First text file. Its lines go in column B, with line numbers in column A. |
| `<fileb>` | Second text file. Its lines go in column C. |
| `--encoding=<name>` | Optional. Encoding of **both** input files. Accepts a name (`windows-1252`, `iso-8859-1`, `utf-16`, …) or a code page number (`1252`). Default: UTF-8, with automatic detection of a byte-order mark (BOM). |
| `-h`, `--help` | Show usage. |

The output file is named `<filea>_<fileb>.xlsx`, using the file names without extension, and is written to the **current directory**. An existing file with the same name is overwritten.

To run from another folder, point `--project` at the project directory:

```
dotnet run --project C:\path\to\SW059_ExcelComparer -- log.txt main.c
```

Or build once and call the compiled program directly:

```
dotnet build -c Release
dotnet bin\Release\net8.0\SW059_ExcelComparer.dll <filea> <fileb>
```

## Output

Given `FileA.txt`:

```
A1
A2
A3
```

and `FileB.txt`:

```
B1
B2
B3
B4
```

running `dotnet run -- FileA.txt FileB.txt` produces `FileA_FileB.xlsx` with a sheet named `Compare`:

| # | FileA.txt | FileB.txt |
|:--|:--|:--|
| 1 | A1 | B1 |
| 2 | A2 | B2 |
| 3 | A3 | B3 |
|   |    | B4 |

- The header row holds the two file names, is bold, and stays visible while scrolling.
- Column `#` holds FileA's line numbers and is blank on rows past the end of FileA.
- When one file is shorter, its column is simply left empty for the remaining rows.
- Lines are stored as plain text exactly as read: a line like `=SUM(A1:A3)` is not turned into a formula, `0042` keeps its leading zeros, and leading spaces are preserved.
- The text uses the monospaced Consolas font, so indentation in source code lines up.

## Limits

- A cell holds at most 32,767 characters. Longer lines are truncated and a warning is printed.
- A sheet holds at most 1,048,575 data rows. Larger files are rejected with an error.

## Exit codes

| Code | Meaning |
|:--|:--|
| 0 | The Excel file was written. |
| 1 | Wrong arguments, unknown encoding, unreadable input file, too many lines, or the output could not be written (for example because it is open in Excel). |

## Try it

The [test](test/) folder contains two sample files: `filea.txt` (120 lines) and `fileb.txt` (80 lines).

```
cd test
dotnet run --project .. -- filea.txt fileb.txt
```

This creates `test\filea_fileb.xlsx`, where rows 81–120 show only the FileA column.
