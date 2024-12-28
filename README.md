# Advanced Power Query Case Study Documentation

This document provides a detailed explanation of the Power Query M code used in this advanced case study. The goal of the transformation is to clean, structure, and enhance data extracted from an Excel file.

## Code Overview
The code processes an Excel file containing raw data, applies multiple transformations, and prepares it for further analysis. Key steps include data type transformations, unpivoting, and pivoting columns, as well as adding calculated fields.

---

## Step-by-Step Explanation

### 1. **Loading the Data**
```m
Source = Excel.Workbook(File.Contents("C:\Users\ntijani001\OneDrive - pwc\Bureau\Sample01.xlsx"), null, true),
input_Sheet = Source{[Item="input",Kind="Sheet"]}[Data]
```
- **Source**: Loads the Excel workbook located at the specified file path.
- **input_Sheet**: Extracts the sheet named `input` for further processing.

### 2. **Defining Data Types**
```m
Source01 = Table.TransformColumnTypes(input_Sheet, {...})
```
- Specifies the data type for each column. Most columns are set to `type any`, while specific columns are set to `type text`.

### 3. **Extracting Headers**
```m
Level1Headers = List.Skip(List.RemoveNulls(Record.ToList(Source01{0})), 4),
Level2Headers = List.LastN(List.Distinct(Record.ToList(Source01{1})), 2),
Level3Headers = List.ReplaceValue(List.FirstN(List.Skip(Record.ToList(Source01{2}), 4), 5), null, "None", Replacer.ReplaceValue)
```
- **Level1Headers**: Extracts and skips the first 4 header values from the first row.
- **Level2Headers**: Retrieves the last 2 distinct header values from the second row.
- **Level3Headers**: Replaces null values with `"None"` in the third row.

### 4. **Skipping Rows**
```m
Data = Table.Skip(Source01, 3)
```
- Removes the first 3 rows of the table, retaining only the data.

### 5. **Transforming Columns**
```m
ConvertColumnsToList = Table.ToColumns(Data),
SkippedTheCommonColumns = List.Skip(ConvertColumnsToList, 4)
```
- **ConvertColumnsToList**: Converts the table into a list of columns.
- **SkippedTheCommonColumns**: Removes the first 4 common columns from the list.

### 6. **Grouping Columns by Date**
```m
GroupedColumnsByDate = List.Transform(List.Split(SkippedTheCommonColumns, 5), each Table.FromColumns(_, {"None", "Quantity", "Revenue", "Quantity1", "Revenue1"}))
```
- Splits columns into groups of 5 and creates tables for each group with predefined column names.

### 7. **Adding Index and Mapping Columns to relevant Dates**
```m
#"Added Index to Map Other Columns" = Table.TransformColumns(#"Converted to Table", {"Column1", each Table.AddIndexColumn(_, "Index")}),
MappedDatesToTables = Table.TransformColumns(#"Added Index", {"Date Index", each Level1Headers{_}})
```
- Adds an index to columns and maps Level1 headers (the dates).

### 8. **Expanding and Cleaning Data**
```m
#"Expanded Column1" = Table.ExpandTableColumn(MappedDatesToTables, ...),
DataCleaned = Table.SelectRows(#"Expanded Column1", each ([None] <> null))
```
- Expands nested tables into rows.
- Removes rows where the `None` column has null values.

### 9. **Joining and Merging Columns**
```m
GettingTheCommonColumns = Table.FromColumns(...),
#"Added Index1" = Table.NestedJoin(...),
#"Expanded Merge" = Table.ExpandTableColumn(...)
```
- Extracts the first 4 columns for common data.
- Performs a nested join and expands the merged data.

### 10. **Reordering and Unpivoting Columns**
```m
#"Reordered Columns" = Table.ReorderColumns(...),
#"Unpivoted Columns" = Table.UnpivotOtherColumns(...)
```
- Adjusts column order.
- Unpivots columns to a long format for easier analysis.

### 11. **Adding Custom Calculations**
```m
#"Added Custom" = Table.AddColumn(..., "Flag", each if Text.Contains([Attribute],"1") then 0 else 1),
Custom1 = Table.TransformColumns(...)
```
- Adds a `Flag` column to differentiate attributes.
- Maps the `Flag` value to Level2 headers(Cash Or Card).

### 12. **Renaming and Pivoting Columns**
```m
#"Renamed Columns1" = Table.RenameColumns(...),
#"Pivoted Column" = Table.Pivot(...)
```
- Renames specific columns.
- Pivots data to a wide format.

### 13. **Final Data Cleanup**
```m
#"Filtered Rows1" = Table.SelectRows(...),
#"Changed Type" = Table.TransformColumnTypes(...),
#"Removed Errors" = Table.RemoveRowsWithErrors(...)
```
- Filters out rows with specific conditions.
- Changes column types to integers.
- Removes rows with errors.

---

## Outputs
- The final table contains cleaned, structured, and ready-to-use data.
- Key columns: `Site`, `Site Name`, `Machine`, `OIC`, `Date`, `Quantity`, `Revenue`, etc.

---

## Key Functions Used
1. **Table.TransformColumnTypes**: Defines column data types.
2. **Table.Skip**: Skips rows.
3. **Table.ToColumns**: Converts tables to column lists.
4. **Table.AddIndexColumn**: Adds index columns.
5. **Table.ExpandTableColumn**: Expands nested tables.
6. **Table.UnpivotOtherColumns**: Converts wide-format data to long-format.
7. **Table.Pivot**: Converts long-format data to wide-format.
8. **Table.RemoveRowsWithErrors**: Removes rows with data errors.

---

## Notes
- Ensure that the file path is accessible and valid.
- Verify column names and data types for consistency.
- This code is optimized for the given dataset structure. Adjustments may be necessary for other datasets.

---
