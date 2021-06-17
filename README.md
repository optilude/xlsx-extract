# xlsx-extract

Extract data from (poorly) structured XLSX files.

This package provides data structures, algorithms and a command line tool to help extract data from Excel files in XLSX format. It builds upon the popular `openpyxl` library. Through the command line, the main usage pattern is:

- Create a "target" workbook. This can contain any tables, formulae or other content you wish. Typically, it would be used to define the format for a report or summary that draws upon one or more "source" workbooks that are not within your control to change (e.g. reports produced by other systems).
- The "target" workbook should contain a sheet called "Config". In this sheet, you will define a number of blocks that describe where to find data in the source workbook(s), and where to place them in the target workbook (i.e. the same workbook that contains the "Config" sheet). Examples of the kind of things you could do include:

    - Define a directory to search files matching a particular pattern (the most recently modified file matching the pattern will be used)
    - Extract a proportion of the filename (e.g. a string representing a date) and use it as a variable later (e.g. to identify a row or column in a table)
    - Locate a sheet in the source workbook matching a pattern
    - Locate a cell or range (table of cells) in the source workbook through rules like "the first cell of the table starts below row 7 and its value matches the following text pattern and the table continues until the next blank row or column", or "the first cell of the table is the first non-blank cell after cell B3 and the table is 6 rows by 8 columns".
    - Locate a cell or table in the target workbook to copy the source value to. Target locators always use simple references (e.g. `"Sheet 1"!A3:B8`, a defined name, or a named table). When working with source or target tables, you can search those tables for rows or columns matching particular rules, such as: "Find the table that starts two rows below the cell with the value "People" and continues until the first blank row and column. Within this table, locate the row that matches the date extracted from the filename, and the column that matches the value "Bob", and copy the cell at this intersecton to the target cell"; or "Find the table named "Table1", and copy it to the target workbook, expanding it by adding rows or columns as needed"; or "Find the table named "Table2", and within it locate the column matching this particular pattern. Copy that column into a table in the target, transposing it so that it lies in the *row* matching a different pattern. Align the values using the labels in the first column of the source table to the first row of the target table, even if they are in a different order."

Note that pattern matches (with the "matches" operator) all use regular expressions. If you are not familiar with these, consult a guide such as https://docs.python.org/3/howto/regex.html#regex-howto or a tool like https://regex101.com.

The next section describes the configuration format in more detail.

To install, use `pip` either in a virtual environment (recommended) or globally:

    $ pip install xlsx-extract

This should install a program called `xlsx-extract`:

    $ xlsx-extract --help

A basic invocation might look like:

    $ xlsx-extract --source-directory /path/to/sources Target.xlsx Report.xlsx

This will read configuration from "Target.xlsx" from the current directory, read its configuration from the "Config" sheet and write the output to the "Report.xlsx" file (note that if you want to write the results back to "Target.xlsx", you should use the `--update` flag rather than specify the same file as target and output). "Target.xlsx" is assumed to identify its own file (via a "File" configuration block), but we specify the directory to search for files with the `--source-directory` argument.

## Excel configuration format

Match and target configuration can be loaded from an Excel sheet in the target workbook, by default called "Config".

The purpose of this sheet is to match (locate) values in a *source* sheet in another workbook (or multiple other workbooks), and define a *target* in the current workbook where those values will be copied to. Configuration is specified in "blocks" of three columns: a parameter (e.g. "File" or "Name"), an operator (e.g. "is", ">", "matches", etc.) and a value (a string, number, date, or boolean). Blocks are separated by at least one blank row. You can have other (typically blank) cells above or before each configuration block, and other content (typically text commentary) in columns to the right of them, but it is recommended to keep the "Config" sheet as simple as possible to avoid confusing the parser.

> Note: You should *not* use formulae to construct configuration blocks. They must be created with simple cell values only.

A directory in which to look for source files and a source file name must be set first. Typically, the directory is set externally (e.g. on the command line) and the file is set with a "File" block at the top of the configuration sheet, though you can also set the directory with a "Directory" block if you want to embed it in the sheet, and you can specify the source file externally (e.g. on the command line) as well. The directory and file must be set one way or the other before any match blocks (which start with "Name"), otherwise the parser won't be able to locate the source data. It is possible to change the file part way through the configuration with a subsequent "File" block, in which case match blocks below that point will use the new file.

The file block consists of a single row: The parameter is "File", the operator is either "is" or "matches", and a file name or pattern. In this documentation, we will refer to that as `"File" | "is" | "My source.xlsx"` representing a row with three column.

If the operator is "is", the filename is matched exactly; if it is "matches" it will be evaluated as a regular expression, which can use a match group (in parentheses) to extract a value. If multiple files match, the most recently modified will always be used. The filename (or matched subset) will be available in the variable `${file}`.

Match/target blocks should start with a block like `"Name" | "is" | "My name"`. The name can be anything but should be unique. It allows us to refer back to the matched value later via a variable. For example, if the name was "Start date" we’d have a variable `${Start date}`.

Variables can be used in any value in match/target blocks after the point of definition. For instance you could use a parameter row like `"Value" | "is" | "Starting on ${Start date}"` and the `${Start date}` placeholder would be replaced by the relevant string.

The other common key is "Sheet", which can take the operator "is" or "matches" and a value to identify the sheet name in the source workbook. If the operator is "matches" you can use a regular expression. If multiple sheets match, the first one will be used. "Sheet" can be skipped if using a "Cell" or "Table" reference (see below) that either includes a sheet name (e.g. "Foo!A3") or refers to a defined name.

Next must come some search criteria.

A single cell can be matched by either reference or value. To match by reference, use the key "Cell" with operator "is" and a value that specifies either a cell reference (e.g. "A3") or an Excel defined name at sheet or workbook level.

To search for a cell value, use the key "Value" with operator "is" (equality), "<", "<=", ">", ">=", "is empty", "is not empty", or "matches" (a regular expression - if it contains a match group in parentheses the matched first group will be the value of the corresponding variable). It is possible to match a cell that is offset from the one found by value (e.g. one column after, or two rows above) by using the keys "Row offset" or "Column offset". These use the operator "is" and take a number which can be positive (after/below) or negative (before/above), e.g. "Row offset" | "is" | -2.

Finally, the search area can be constrained by using the keys "Min row", "Min column", "Max row" and "Max column", all of which use the operator "is" and take a numeric value (column letters are also allows for min/max columns).

Matching a range or table allows you to populate a table or search for a cell by row and column label.

The easiest way to match a table is by giving a "Table" that "is" a range (e.g. "A3:C4") or a defined name specifying a range, or a named Excel table.

If that’s not possible (i.e. the structure of the source workbook is not sufficiently stable), you can search for it by finding the start cell (top/left) and giving one of an end cell (bottom/right); a fixed number of rows and columns; or assuming that the table extends across the contiguous width of the header row and height of the first column (i.e. until a blank cell is found).

Searching for the start cell uses a cell match as described above, but each key is prefixed by the word "Start", so it can be identified by "Start cell" or "Start value", optionally with "Start row offset" or "Start column offset". You do not need specify a sheet or search area for the start cell specifically (i.e. there are no "Start sheet" or "Start min row" parameters) as these are set at the match block level.

If you specify only a start cell, the table dimensions will be found by looking for a contiguous row of header cells and a contiguous first column, until a blank cell is found (empty cells are allowed within the table).

To specify an end cell, use the same approach with the prefix "End", e.g. "End cell" or "End value".

To specify a fixed size, use the keys "Rows" and "Columns" with the operator "is" and a number for each.

As with single cells, the search area can be constrained using "Min row", "Max row", "Min column" and "Max column". These are specified at thet top level, i.e. not for the start or end cell if used.

Once we have matched the cell, the next step is to identify the target. The target is actually optional, because it can be useful to match a cell and then use its value as a variable in another match, but most blocks will specify a target, using a row like `"Target cell" | "is" | "<reference>"`. Targets can only be specified by a cell/range reference (`"Sheet 1!A3"`), a defined name at workbook level (i.e. not a sheet-specific name), or as a named Excel table. The assumption is that you have enough control over the target workbook to not need to search by value or offset. Use "Target cell" for cells and "Target table" for ranges.

The simplest scenario is where you match a single cell and copy it to a single target. Simplify specify a single-cell target range or name, and the the matched value will be copied over.

Alternatively, you can copy a whole matched range into a target range. Source and target should have the same number of columns. Rows will be replaced and reduced/expanded as needed. If you would rather keep the height of the target table intact (leaving unused rows untouched/truncating the source table), add a parameter `"Expand" | "is" | "FALSE"`.

You can also copy a single cell from a source table into a single target cell. In this case, you need to specify which row/column intersection to look in. This is done using the keys "Source row value" and "Source column value". These behave like the "Value" parameters for a single cell match (i.e. they can use any operator, including regular expression matching), and are constrained to search the first row and first column of the matched table, respectively. You can use "Source row offset" and "Source column offset" as well.

Finally, you can copy an entire row or column from the source into a specified row or column in a target table. This requires that the target is a range (or named table).

First, you specify one of "Source row value" or "Source column value"; and one of "Target row value" or "Target column value" (offsets are also possible for each, e.g. "Source column offset" | "is" | 1). If you specify a source row and a target column or vice-versa, the data will be transposed.

If targetting a row, the "Expand" parameter, defaulting to TRUE, behaves a little differently from the scenario of copying an entire range: it will only expand the table, but not delete rows. It does nothing if targeting a column.

Instead of copying the entire matched row/column, however, you can align it to the corresponding column/row headings by setting "Align" | "is" | TRUE. For example, imagine a source table with rows where the first column contains a set of labels like "Alpha", "Beta", "Delta", "Gamma". In the target table, you might have some of these labels, perhaps in a different order, e.g. "Gamma", "Delta". If you target a particular source column and target column and set "Align" to TRUE, the values corresponding to "Gamma" and "Delta" (only) will be copied to the relevant rows in the target column.

## Limitations

This tool cannot handle every conceivable scenario, and you may need to be a little creative in how to construct the target workbook. Some of these limitations come from the way that the source and target workbooks are parsed and manipulated using the `openpyxl` Python library.

- Each *source* workbook is loaded in "data only" mode. This means that the values that will be read will be those saved with the workbook when it was last openened in Excel. Thus, dynamic formulae (e.g. using the current date or attempting to read data from an external source) will not be re-evaluated.
- The *target* workbook, conversely, is opened in "formulas" mode. This means that `xlsx-extract` will see the text of a formula, but it is *not* able to evaluate it. This is only really a problem if you want to use formulas on the "Config" sheet (which, put simply, you can't).
- When replacing tables in the target workbook and using the "Expand" parameter, you may change the shape of the target workbook by adding or removing rows or columns. It is possible for this to corrupt the workbook if adjacent cells or named tables are impacted. In general, it is safest to keep tables that may grow or shrink on their own, simple worksheets and reference this data from other sheets as required.
- Similarly, formulae elsewhere on a sheet are not updated if the geometry of the sheet changes. In Excel, if you add or delete a row or column, any formulae that reference cells that change coordinates as a result are automatically updated. Sadly, this tool cannot do that for you. Again, keeping growing/shrinking tables on their own sheets can help with this.
- It is possible for other, more advanced elements of the target workbook to be stripped when the output file is saved.

## Using the Python interface

All of the above can be done using primities in Python. The code contains the documentation, but for orientation:

- `xlsx_extract.range.Range` is a class that defines one or more cells. It is a thin wrapper around the OpenPyXL cell API, but provides helpful functions for inspecting and working with ranges. Matches return `Range` instances.
- `xlsx_extract.match` defines the parametrs that correspond to match blocks in the Excel configuration. There are two flavours, `CellMatch` and `RangeMatch` which operate on single cells and ranges (tables), respectively. The main method is `match()` which operates on a workbook to return a matched range and key value (which is what gets stored in variables in the Excel configuration).
- `xlsx_extract.target.Target` defines parameters to identify a target location. The `extract()` method operates on a source match (a `CellMatch` or `RangeMatch`), a target match (ditto) and source and target workbooks to extract data from the source and place it in the target.

Take a look at the various test files (e.g. `range_test.py`, `match_test.py` and `target_test.py`) for examples of how to invoke these.