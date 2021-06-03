# xlsx-extract

Extract data from (poorly) structured XLSX files.

## Excel configuration data

Match and target configuration can be loaded from an Excel sheet in the target workbook, by default called "Config".

The purpose of this sheet is to match (locate) values in a "source" sheet in another workbook (or multiple other workbooks), and define a "target" in the current workbook where those values will be copied to.						
The configuration starts with a row containing three adjacent cells: "File", either "is" or "matches", and a file name or pattern. If the operator is "is", the filename is matched exactly; if it is "matches" it will be evaluated as a regular expression, which can use a match group (in parentheses) to extract a value. If multiple files match, the most recently modified will always be used. The filename (or matched subset) will be available in the variable ${file}.

Next come one or more match/target blocks. All blocks after this will match content of the file, but you can repeat the "File" line to target a new file in subsequent blocks.

Match/target blocks are separated by blank rows and have three columns: key, operator, and value. The first row should have a key "Name", the operator "is", and a unique name in the value column. This uniquely identifies the match, and allows us to refer back to the matched value later via a variable. For example, if the name was "Start date" we’d have a variable ${Start date}.

Variables can be used in any value in match/target blocks after the point of definition. For instance you could use a parameter row like "Value" | "is" | "Starting on ${Start date} " and the ${Start date} placeholder would be replaced by the relevant string.

The other mandatory key is "Sheet", which can take the operator "is" or "matches" and a value to identify the sheet name in the source workbook. If the operator is "matches" you can use a regular expression. If multiple sheets match, the first one will be used.

Next must come some search criteria.

A single cell can be matched by either reference or value. To match by reference, use the key "Cell" with operator "is" and a value that specifies either a cell reference (e.g. "A3") or an Excel defined name at sheet or workbook level.

To search for a cell value, use the key "Value" with operators "is" (equality), "<", "<=", ">", ">=", "is empty", "is not empty", or "matches" (a regular expression). It is possible to match a cell that is offset from the one found by value (e.g. one column after, or two rows above) by using the keys "Row offset" or "Column offset". These use the operator "is" and take a number which can be positive (after/below) or negative (before/above), e.g. "Row offset" | "is" | -2.

Finally, the search area can be constrained by using the keys "Min row", "Min column", "Max row" and "Max column", all of which use the operator "is" and take a numeric value (column letters are also allows for min/max columns).

Matching a range or table allows you to populate a table or search for a cell by row and column label.

The easiest way to match a table is by giving a "Table" that "is" a range (e.g. "A3:C4") or a defined name specifying a range, or a named Excel table.

If that’s not possible (i.e. the structure of the source workbook is not sufficiently stable), you can search for it by finding the start cell (top/left) and giving one of an end cell (bottom/right); a fixed number of rows and columns; or assuming that the table extends across the contiguous width of the header row and height of the first column (i.e. until a blank cell is found).

Searching for the start cell uses a cell match as described above, but each key is prefixed by the word "Start", so it can be identified by "Start cell" or "Start value", optionally with "Start row offset" or "Start column offset". Do not specify a sheet or search area for the start cell specifically (i.e. there are no "Start sheet" or "Start min row" parameters).

If you specify only a start cell, the table dimensions will be found by looking for a contiguous row of header cells and a contiguous first column (empty cells are allowed within the table).

To specify an end cell, use the same approach with the prefix "End", e.g. "End cell" or "End value".

To specify a fixed size, use the keys "Rows" and "Columns" with the operator "is" and a number for each.

As with single cells, the search area can be constrained using "Min row", "Max row", "Min column" and "Max column". These are specified at thet top level, i.e. not for the start or end cell if used.

Once we have matched the cell, the next step is to identify the target. The target is actually optional, because it can be useful to match a cell and then use its value as a variable in another match, but most blocks will specify a target, using a row like "Target cell" | "is" | "<reference>". Targets can only be specified by a cell/range reference ("A3" or "A3:C7"), a defined name at workbook level (i.e. not a sheet-specific name), or as a named Excel table. The assumption is that you have enough control over the target workbook to not need to search by value or offset. Use "Target cell" for cells and "Target table" for ranges.

The simplest scenario is where you match a single cell and copy it to a single target. Simplify specify a single-cell target range or name, and the the matched value will be copied over.

Alternatively, you can copy a whole matched range into a target range. Source and target should have the same number of columns. Rows will be replaced and reduced/expanded as needed. If you would rather keep the height of the target table intact (leaving unused rows untouched/truncating the source table), add a parameter "Expand" | "is" | "FALSE".

You can also copy a single cell from a source table into a single target cell. In this case, you need to specify which row/column intersection to look in. This is done using the keys "Source row value" and "Source column value". These behave like the "Value" parameters for a single cell match (i.e. they can use any operator, including regular expression matching), and are constrained to search the first row and first column of the matched table, respectively. You can use "Source row offset" and "Source column offset" as well.

Finally, you can copy an entire row or column from the source into a specified row or column in a target table. This requires that the target is a range (or named table).

First, you specify one of "Source row value" or "Source column value"; and one of "Target row value" or "Target column value" (offsets are also possible for each, e.g. "Source column offset" | "is" | 1). If you specify a source row and a target column or vice-versa, the data will be transposed.

If targetting a row, the "Expand" parameter, defaulting to TRUE, behaves a little differently from the scenario of copying an entire range: it will only expand the table, but not delete rows. It does nothing if targeting a column.

Instead of copying the entire matched row/column, however, you can align it to the corresponding column/row headings by setting "Align" | "is" | TRUE. For example, imagine a source table with rows where the first column contains a set of labels like "Alpha", "Beta", "Delta", "Gamma". In the target table, you might have some of these labels, perhaps in a different order, e.g. "Gamma", "Delta". If you target a particular source column and target column and set "Align" to TRUE, the values corresponding to "Gamma" and "Delta" (only) will be copied to the relevant rows in the target column.