---
prev: /sql/orgcharts/
next: /views/
---

# SQL Syntax

**Structured Query Language (SQL)** is one of the primary languages used to retrieve and manipulate data in relational systems. Microsoft Excel is often used in a similar fashion, storing information in worksheets that resemble database tables: columns with headers and many rows of data beneath them.

Excel supports SQL through its own dialect. Earlier versions relied on the Jet engine developed for Microsoft Access, while modern versions (beginning with Office 2007 and the `.xlsx` file format) use the ACE engine. Neither engine fully conforms to ANSI SQL standards, and both include unique behaviors and limitations that differ from traditional database systems.

This page consolidates information from many [sources](#resourcesacknowlegements) to provide a clear reference for the SQL syntax and functions supported by Microsoft Excel’s SQL engine.

## Key Concepts
### Tables

Tables are represented by Excel worksheets. A table is specified as the worksheet name followed by a `$` character. The square bracket characters `[` and `]` are used as delimiters, allowing worksheet names that contain spaces or special characters to be referenced safely.

In Excel SQL, table names follow this pattern:

_Syntax:_

```sql
    [worksheetname$]
```

_Example:_

```sql
    [Past Invoices$]
```

The example above refers to an open-ended Worksheet named `Past Invoices`. 

It is also possible to use a portion of a worksheet, and not the whole worksheet. To use a partial worksheet, specify a cell range using standard Excel range notation. For example:

_Example:_

```sql
    [Past Invoices$B10:G700]
```

Refers to to range of cells `B10:G700` within the _Past Invoices_ Worksheet.

---

### Columns

Column names can refer to the Excel heading for the column, such as A, B, C or the column heading in the first row of the table (or table range). Column names are limited to 64 characters. Column names containing blanks should also be enclosed in square brackets as in the following example. It is generally a good habit to always delimit column names with the square brackets.

_Example:_

```sql
    [Customer Name]
```

The data type of a column is determined by the ACE engine by scanning the values in the first 8 rows in that column. If the scanned rows are blank then ACE assumes a String data type. If a column is determined to be numeric, then any rows with alpha characters in this column will be returned as Null.

---

### Aliases

Aliases are used to temporarily assign a different name to a table or column heading using an `AS` clause. Basically aliases are created to make column names more readable. The `WHERE`, `ORDER BY`, `GROUP BY`, `HAVING` and `JOIN` clauses support aliases.

SQL alias syntax for columns follows the syntax format:

_Syntax:_

```sql
    SELECT [column_name] AS column_alias_name FROM [worksheet$]
```

SQL alias syntax for tables follows the syntax format:

_Syntax:_

```sql
    SELECT [column_name] FROM [worksheet$] AS table_alias_name
```

The following example shows multiple ways in which an alias can be expressed within an Excel SQL statement.

_Example:_

```sql
    SELECT [country] AS CC, [Units Sold]
    FROM [Sales$] AS sales
    WHERE [sales.segment] = 'Government' AND [sales].[CC] = 'US'
```

---

### Strings
Most SQL statements involve the use of string values as criteria in a [WHERE clause](#where-clause), or will return string values in the column results of a [SELECT clause](#select-clause). 

Strings are denoted in SQL using the single quotation mark `'` character, as opposed to the double quote `"` character used in VBA and other programming languages. A String value such as `'Central Region'` is properly formatted for use as selection criteria, or as a value to be inserted.

Strings can be concatenated using the ampersand `&` character. For example, the statement  `'Central' & ' ' & 'Region'` equates to the String value `'Central Region'`. 

Microsoft Excel SQL provides numerous String functions that perform operations on an input String and return a String or numeric value result. These functions can be combined to create unique input or output values depending upon the SQL statement. A table of commonly used [String Functions](#string-functions) is provided at the end of this document as reference.

---

### Null Values
A Null Value is a value that is unavailable, unassigned, unknown or inapplicable. If a row lacks the data value for a particular column, that value is said to be null or to contain a null. A Null Value is like a character string of length zero, however you should never use a Null Value to represent a value of zero. 

Since Microsoft Excel SQL allows null values, it is your responsibility to take them into consideration in your SQL statements. [Null Handling](#null-handling) is described in greater detail later in this document.

## Querying Data

Microsoft Excel SQL queries support `FROM`, `WHERE`, `GROUP BY`, `HAVING`, `ORDER BY`, and `JOIN` clauses. Excel SQL `SELECT` statements conform to the following pattern:

```sql
    SELECT [DISTINCT] [TOP [PERCENT] n] * | column1, column2 [AS alias2], ...
    [FROM worksheet]
    [WHERE condition]
    [GROUP BY column1, column2, ...]
    [HAVING condition]
    [ORDER BY column1 [DESC|ASC], column2 [DESC|ASC], ...]
```

## SELECT clause

The `SELECT` clause is used to execute a query to select data from a Microsoft Excel worksheet.

---

### `SELECT *`

A query that selects all rows and columns from the Excel file.

_Example:_

```sql
    SELECT * FROM [Sales$];
```

In this example, the query fetches all rows and columns in the Sales worksheet. You can query against different sheets in a common Excel file using this syntax.

---

### `SELECT TOP n`

The `TOP` option specifies how many rows to return.

_Example:_

```sql
    SELECT TOP 5 [Quantity], [Name], [Price] FROM [Sales$];
```

Here a maximum of 5 rows will be returned.

---

### `SELECT TOP n PERCENT`

The `TOP PERCENT` option specifies the percentage of the complete result set to return.

_Example:_

```sql
    SELECT TOP 25 PERCENT [Quantity], [Name], [Price] FROM [Sales$];
```

Here the first 25% of the result set rows will be returned. If the query result set contains 400 rows, then the top 100 rows will be returned.

---

### `SELECT` column list

Create a query that selects specific columns from the Excel file.

_Example:_

```sql
    SELECT [Quantity], [Name], [Price] FROM [Sales$];
```

---

### `SELECT` from a range of cells

Limit your query to a specific cell range.

_Example:_

```sql
    SELECT [Quantity], [Name], [Price] FROM [Sales$A1:E101];
```

In this example, we do not impose any limitations on the values themselves. However, we direct the query to look only at a Range of cells (A1 through E101). Note that the cell range is specified after the dollar sign in the table name, using the colon between the first cell and the final cell in the range.

---

### `SELECT DISTINCT`

A column can contain duplicate values, and to list the distinct values, use the `SELECT DISTINCT` clause. The `DISTINCT` clause can be used to return only distinct values from a set of records.

_Example:_

```sql
    SELECT DISTINCT [Name], [Price] FROM [Sales$A1:E101];
```

---

### Aggregate functions

An aggregate function performs a calculation on a set of values, and returns a single value. Except for `COUNT(*)`, aggregate functions ignore null values. Aggregate functions are often used with the `GROUP BY` clause of the `SELECT` statement.

|  Function  | Description                                                                                                                                                   |
| :--------: | ------------------------------------------------------------------------------------------------------------------------------------------------------------- |
|   `AVG`    | Returns the average (mean) value. It ignores null values. If no rows are selected, the result is NULL.                                                        |
| `COUNT(*)` | Returns the number of rows in a specified table, and it preserves duplicate rows. It counts each row separately. This includes rows that contain null values. |
|   `MAX`    | Returns the highest value.                                                                                                                                    |
|   `MIN`    | Returns the lowest value.                                                                                                                                     |
|  `STDEV`   | Returns the statistical standard deviation of all values in the specified expression.                                                                         |
|  `STDEVP`  | Returns the statistical standard deviation for the population for all values in the specified expression.                                                     |
|   `SUM`    | Returns the sum of all values.                                                                                                                                |
|   `VAR`    | Returns the statistical variance of all values in the specified expression.                                                                                   |
|   `VARP`   | Returns the statistical variance for the population for all values in the specified expression.                                                               |

_Example:_

```sql
    SELECT COUNT(*)
    FROM [Sales$];
```

In this example the number of rows in the Sales worksheet is returned.

_Example:_

```sql
    SELECT SUM([Units Sold]), AVG([Units Sold])
    FROM [Sales$]
```

In this example the total number of units sold, and the average quantity of units sold is returned.

---

### Arithmetic

It is possible within a `SELECT` statement to run mathematical operations on two expressions of one or more data types.

#### Arithmetic Operators

Arithmetic operators can be used to perform mathematical operations against query results. They are:

| Operator | Description                                         |
| :------: | --------------------------------------------------- |
|    \+    | Addition                                            |
|    \-    | Subtraction                                         |
|    \*    | Multiplication                                      |
|    \/    | Division                                            |
|   MOD    | Modulo, returns the integer remainder of a division |

_Example:_

```sql
    SELECT [units sold] MOD 10 AS UNITS_SOLD_MOD_10
    FROM [Sales$]
```

Here the number of units sold is calculated as modulo 10 and the result is expressed using the alias `UNITS_SOLD_MOD_10`.

_Example:_

```sql
    SELECT [sale price], ([sale price] * 1.06) AS [sale price plus tax]
    FROM [Sales$]
```

In this example the sale price is multiplied by 1.06 to reflect a 6% sales tax, and is expressed using the alias sale price plus tax.

#### Mathematical Functions

The following scalar functions perform a calculation, usually based on input values that are provided as arguments, and return a numeric value:
| Function | Description |
| :------: | --------------------------------------------------- |
| `ABS` | A mathematical function that returns the absolute (positive) value of the specified numeric expression. (ABS changes negative values to positive values. ABS has no effect on zero or positive values.) |
| `COS` | A mathematical function that returns the trigonometric cosine of the specified angle - measured in radians - in the specified expression. |
| `EXP` | Returns the exponential value of the specified expression. |
| `LOG` | Returns the natural logarithm of the specified expression. |
| `ROUND` | Returns a numeric value, rounded to the specified length or precision. |
| `SIN` | Returns the trigonometric sine of the specified angle, in radians, and in an approximate numeric, float, expression. |
| `TAN` | Returns the tangent of the input expression. |

_Example:_

```sql
    SELECT EXP(LOG(20)), LOG(EXP(20))
```

Here the SQL statement returns the exponential value of the natural logarithm of 20 and the natural logarithm of the exponential of 20. Because these functions are inverse functions of one another, the return value in both cases is 20.

### Conditionals

Conditionals are used to conditionally modify results. In Excel SQL, this is done with the `IIF()` function. The function signature is: `IIF(expression, truepart, falsepart)`. If expr evaluates to true, then truepart is returned, otherwise, falsepart is returned.

_Example:_

```sql
    SELECT [Product], [Units Sold], [Sale Price],
    IIF([Sale Price] < 10.00,'SPECIAL!','')
    FROM [Sales$]
```

Here the SQL returns the string 'SPECIAL!' if the price is less than $10.00; otherwise it returns a blank string.

### Null Handling

Under some conditions, such as when a cell is has no data, Excel returns a Null value. You can test whether a cell is null using the `IsNull(expression)` function. `IsNull()` returns a Boolean true value (=-1) if the argument is null, and a Boolean false (=0) if the argument is not null.

A simple example is:

_Example:_

```sql
    SELECT IsNull([Units Sold])
    FROM [Sales$]
```

The `IsNull()` function can be combined with the `IIF()` syntax to return a specific value in cases where a null is found in a column, and the actual column value where the value is not null.

### Aggregate Functions + Conditionals + Null Handling

Conditional tests are very useful in situations where you are aggregating data which may have `NULL` values. You can ensure predictable behavior by using `IFF` to return a consistent value within an aggregate function when `NULL` is encountered. For example, if you want `NULL` to be treated a zero, you write the SQL as:

_Example:_

```sql
    SELECT MIN(IIf(IsNull([Units Sold]), 0, [Units Sold]))
    FROM [Sales$]
    WHERE [country] = 'Canada'
```

### Formatting date values

#### `FORMATDATETIME` function

The `FORMATDATETIME` function has a set of fixed options for formatting, shown below.
|Format Code|Description|
|:---------:|-----------|
|0|Display a date and/or time. Date parts are displayed in short date format. Time parts are displayed in long time format. For example, "1/1/2014"|
|1|Display a date using the long date format specified in your computer's regional settings. For example, "Wednesday, January 1, 2014"|
|2|Display a date using the short date format specified in your computer's regional settings. For example, "1/1/2014".|
|3|Display a time using the time format specified in your computer's regional settings. For example "12:00:00 AM".|
|4|Display a time using the 24-hour format (hh:mm). For example, "00:00".|

_Example:_

```sql
    SELECT
    [Date],
    FORMATDATETIME([Date],0) AS [FormatDateTime0],
    FORMATDATETIME([Date],1) AS [FormatDateTime1],
    FORMATDATETIME([Date],2) AS [FormatDateTime2],
    FORMATDATETIME([Date],3) AS [FormatDateTime3],
    FORMATDATETIME([Date],4) AS [FormatDateTime4]
    FROM [Sales$]
```

#### `FORMAT` function

If you need a more flexible set of formatting options, the `FORMAT` function takes a format template string:
|String|Description|
|:----:|-----------|
|d|display the day as a number without a leading zero (1 - 31).|
|dd|Display the day as a number with a leading zero (01 - 31).|
|ddd|Display the day as an abbreviation (Sun - Sat).|
|m|Display the month as a number without a leading zero (1 - 12). If m immediately follows h or hh, the minute rather than the month is displayed.||
|mm|Display the month as a number with a leading zero (01 - 12). If m immediately follows h or |hh, the minute rather than the month is displayed.|
|mmm|Display the month as an abbreviation (Jan - Dec).|
|mmmm|Display the month as a full month name (January - December).|
|oooo|The same as mmmm, only it's the localized version of the string.|
|y|Display the day of the year as a number (1 - 366).|
|yy|Display the year as a 2-digit number (00 - 99).|
|yyyy|Display the year as a 4-digit number (100 - 9999).|

_Example:_

```sql
    SELECT [Date], FORMAT([Date],"yyyy-mmm-dd") AS [formatted_date],
    FROM [Sales$]
```

In the example above, `[Date]` will be returned as an absolute number such as 41640, and `[formatted_date]` will be returned as a string such as 2014-Jan-01 as specified by the format pattern.

### Formatting numbers

The `FORMAT` function can also be used to format date formats other than dates. Assume we have a query which calculates sales tax as 6% of the sale price. The query would look as follows:

_Example:_

```sql
    SELECT
      [Gross Sales],
      [Gross Sales]*0.06 as [Sales Tax]
    FROM [Sales$]
```

The results of this query would have varying decimal places of 0 or more. To always return 2 decimal places we can modify the query as follows:

_Example:_

```sql
    SELECT
      [Gross Sales],
      FORMAT([Gross Sales]*0.06,'0.00') as [Sales Tax]
    FROM [Sales$]
```

The `[Sales Tax]` column is formatted to two decimal places. Note that the formatting code must be enclosed in single quotes, not double quotes. Note also that the format function returns a string type, so if you want to do math with a formatted value, you’ll have to later convert it to a numerical type with a type conversion function.

An alternate way to format numbers is to use the named numeric formats. Since we were formatting the sales tax with two decimal places to represent currency, we could also use the 'Currency' named format. This format will display the number with the currency symbol, the thousand separator, and if appropriate; display two digits to the right of the decimal separator. Output is based on system locale settings.

_Example:_

```sql
    SELECT
      [Gross Sales],
      FORMAT([Gross Sales]*0.06,'Currency') as [Sales Tax]
    FROM [Sales$]
```

Here the `[Sales Tax]` will be returned with an appearance such as $2,224.80.

For more information on the `FORMAT` function and a listing of format options for Numbers, Dates and times, Date and time serial numbers, and Strings, please refer to [Microsoft's documentation of the FORMAT function](https://docs.microsoft.com/en-us/office/vba/Language/Reference/User-Interface-Help/format-function-visual-basic-for-applications).

## FROM clause

The `FROM` clause determines the specific dataset to examine to retrieve data. For Excel, this dataset would be indicated as a specific worksheet, with a `$` appended to the worksheet name.

_Example:_

```sql
    SELECT * FROM [Sales$];
```

In this example the Sales worksheet is specified as `[Sales$]`

## WHERE clause

The `WHERE` clause, sometimes called the predicate, states the qualifying conditions for a query. You can combine arithmetic operators and the comparison operators in the `WHERE` clause. Multiple conditions can be joined by the `AND` and `OR` clauses, optionally surrounded by (parentheses) to group them. Only the records that satisfy the specified criteria are returned by the query.

### `WHERE` comparison operators

Comparison operators test whether two expressions are the same. The following list contains the comparison operators supported in a `WHERE` clause. `WHERE` conditions follow the syntax:

`WHERE` column `operator` predicate

|   Operator    | Description                                                                                                                                                                           |
| :-----------: | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
|      `=`      | Equals operator. Compares the equality of two expressions.                                                                                                                            |
|     `!=`      | Not Equal To operator. Compares two expressions and returns TRUE if the left operand is not equal to the right operand; otherwise, the result is FALSE.                               |
|     `<>`      | Not Equal To operator. Compares two expressions and returns TRUE if the left operand is not equal to the right operand; otherwise, the result is FALSE.                               |
|      `>`      | Greater Than operator. Compares two expressions and returns TRUE if the left operand has a value higher than the right operand; otherwise, the result is FALSE.                       |
|     `>=`      | Greater Than or Equal To operator. Compares two expressions and returns TRUE if the left operand has a value greater than or equal to the right operand; otherwise, it returns FALSE. |
|     `<=`      | Less Than or Equal To operator. Compares two expressions and returns TRUE if the left operand has a value lower than or equal to the right operand; otherwise, it returns FALSE.      |
|      `<`      | Less Than operator. Compares two expressions and returns TRUE if the left operand has a value lower than the right operand; otherwise, the result is FALSE.                           |
|   `IS NULL`   | A Null value is a value that is unavailable, unassigned, unknown or inapplicable. This operator is used to determine if a field doesn't contain data.                                 |
| `IS NOT NULL` | This operator is used to determine of a field does contain data.                                                                                                                      |

When specifying the condition, value must be an exact match of the column value in the worksheet.

String value comparisons are case sensitive. You can use the `UCASE()` (i.e. upper case) or `LCASE()` (i.e. lower case) function to perform case insensitive comparisons.




### `WHERE` logical operators

Comparison operators test whether two expressions are the same. The following list contains the logical operators supported in a `WHERE` clause.

| Operator | Description                                                                                                                                                                                                            |
| :------: | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
|   `ALL`    | Returns TRUE when all of the subquery values meet the condition.                                                                                                                                                       |
|   `AND`    | Combines two Boolean expressions and returns TRUE when both expressions are TRUE.<br><br>_Syntax:_<br> condition `AND` condition                                                                                       |
|   `ANY`    | Returns TRUE when any of the subquery values meet the condition.<br><br>_Syntax:_<br>condition `ANY` condition                                                                                                         |
| `BETWEEN`  | Specifies a range to test. Returns TRUE when the operand is within the range of comparisons.<br><br>_Syntax:_<br>column `BETWEEN` value1 `AND` value2                                                                  |
|  `EXISTS`  | Specifies a subquery to test for the existence of rows. Returns TRUE when the subquery returns one or more records.<br><br>_Syntax:_<br>column `EXISTS`                                                                |
|    `IN`    | Determines whether a specified value matches any value in a subquery or a list. Returns TRUE when the operand is equal to one of a list of expressions.<br><br>_Syntax:_<br>column `IN` (value [, value ...])          |
|   `LIKE`   | Determines whether a specific character string matches a specified pattern. Returns TRUE when the operand matches a pattern.<br><br>_Syntax:_<br>column `LIKE` like_expression                                         |
|   `NOT`    | Negates a Boolean input (it reverses the value of any Boolean expression). It therefore returns TRUE when the expression is FALSE.<br><br>_Syntax:_<br>column `NOT EXISTS`<br>column `NOT BETWEEN` value1 `AND` value2 |
|    `OR`    | Combines two conditions. Returns TRUE when either of the conditions is TRUE.<br><br>_Syntax:_<br> condition `OR` condition                                                                                             |
|   `SOME`   | Same as ANY. Returns TRUE when any of the subquery values meet the condition.                                                                                                                                          |

### Boolean Precedence

When using `AND` and `OR` to specify multiple conditions, use (parentheses) to group the conditions. If no parentheses are specified, the conditions specified with `AND` are evaluated together.

_Syntax:_

The expression

```sql
    condition1 AND condition2 AND condition3 OR condition4
```

is equivalent to

```
    (condition1) AND (condition2) AND (condition3 OR condition4)
```

_Syntax:_

The expression

```
    condition1 AND condition2 OR condition3 AND condition4
```

is equivalent to

_Syntax:_

```
    (condition1) AND (condition2 OR condition3) AND (condition4)
```

### `WHERE` with multiple criteria

Use a `WHERE` clause in your query to filter your Excel data.

_Example:_

```sql
    SELECT [Quantity], [Name], [Price] FROM [Sales$]
    WHERE [Sale ID] >= 23 AND [Sale ID] <= 28;
```

In this example, we limit our result set to records whose `[Sale ID]` is >= 23 and < 28. The syntax for column names in the `WHERE` clause uses square brackets, as we saw previously

### `WHERE` with string values

According to the SQL standard, the text delimiter in SQL is the single quotation mark (`'`). Use the single quote character (`'`) in your query to denote strings to match against your Excel data.

_Example:_

```sql
    SELECT * FROM [Sales$]
    WHERE [Country] = 'Canada'
```

In this example, we limit our result set to records where the `[Country]` value matchs the string Canada.

[String Functions](#string-functions) allow you to create a query using a string function in the `WHERE` clause, or to modify the column results of the query.

_Example:_

```sql
    SELECT UCASE([Name]),[Price]
    FROM [Sales$]
```

In the example above, the VBA `UCASE()` function is being used to convert the `NAME` values to upper case.

_Example:_

```
    SELECT [Sale ID], [Sale Date], [Quantity], [Name], [Price]
    FROM [Sales$]
    WHERE MID([Name],1,4) = 'NYNY'
```

Here the VBA `MID()` function is being used to limit our results to the records which contain the string 'NYNY' in the first four characters of the `NAME` column.

_Example:_

```sql
    SELECT [Name], [Price], 'Central' AS [REGION]
    FROM [Sales$]
```

Here a constant value of 'Central' will be included with every row returned under the column name [REGION].

In Excel SQL the ampersand `&` symbol is used to peform string concatenation. In the example below the table values of `[Street]`, `[City]`, `[State]`, and `[Zip Code]` will be concatentated along with formatting strings to build a single address string returned as the column name `[Address]`.

_Example:_

```sql
    SELECT
    [Street] & ', ' & [City], & ' ' & [State] & ' ' & [Zip Code] & ' USA' AS [Address]
    FROM [Sales$]
```

### `WHERE` with date values

Date literals need to be enclosed in delimiters in SQL strings. The hash sign `#` acts as the delimiter for date values. Use the `#` character in your query to filter your Excel data by date.

_Example:_

```sql
    SELECT * FROM [Sales$]
    WHERE [Date] BETWEEN #9/1/2013# AND #12/31/2013#
```

In this example, we limit our result set to records between the dates of September 1, 2013 and December 31, 2013 by sepcifying the date in the local format and wrapping the value with `#` characters so a date comparison is performed against the date serial number instead of treating the date as a string value.

### `WHERE` with delimiters in text literals

Consider the case where you have a text field in your table and you want to query to query it within the WHERE condition of a query. In the query the criteria values themselves, however, contain the text delimiter, i.e. the single quote. This would be the case when searching for the name O'Brien. If you a text literal of 'O'Brien' inside your SQL statement it will cause a syntax error. The solution is to double the single quote inside the text literal and it will be treated as just one single quote inside the literal.

_Example:_

```sql
    SELECT * FROM [Sales$] WHERE [CustomerName] = 'Martha O''Brian'
```

## GROUP BY clause

The `GROUP BY` clause is used in combination with aggregate functions (e.g. SUM) to group the results by one or more columns.

_Example:_

```sql
    SELECT DISTINCT [Name], SUM([Quantity])
    FROM [Sales$]
    GROUP BY [Name]
```

`GROUP BY` must include ALL the non-aggregated `SELECT` expressions. The first column is not enough.

## HAVING clause

The `HAVING` clause states the qualifying conditions for aggregated values. It is used in conjunction with aggregate functions to filter aggregated values.

A typical select statement using the `HAVING` clause will follow the syntax pattern:

_Syntax:_

```sql
    SELECT [column1], aggregate_function([column2]) FROM [worksheet$]
    GROUP BY [column1]
    HAVING aggregate_function([column2]) comparison_operator value
```

as in the following example which returns the count of the number of orders associated with a name if the count is greater than 25.

_Example:_

```sql
    SELECT [name], COUNT([orders])
    FROM [Sales$]
    GROUP BY [name]
    HAVING COUNT([orders]) > 25
```

## ORDER BY clause

`ORDER BY` is used to sort the results by one or more columns and sorts in ascending order by default. To sort in descending order, use the `DESC` keyword.

A typical SELECT statement using the ORDER BY clause will follow the pattern:

_Syntax:_

```sql
    SELECT [column1], [column2]
    FROM [worksheet$]
    ORDER BY [column1] ASC | DESC, [column2] ASC | DESC
```

where `ASC` refers to Ascending Order (A-Z, 0-9), and `DESC` refers to Descending Order (Z-A, 9-0). For example:

_Example:_

```sql
    SELECT *
    FROM [Sales$]
    ORDER BY [units sold] DESC
```

In this example all the rows and columns from the `Sales` worksheet are selected and returned in descending order as determined through the number of units sold.

The `ORDER BY` clause cannot be used on a column with mixed data types. If a column contains data with mixed data types, the data needs to be converted to one data type.

Note that the `ORDER BY` clause also cannot reference aliases, as you will receive the error `"No value given for one or more required parameters"`. Instead, use the ordinal index in the `SELECT` clause as follows to order the results:

_Example:_

```sql
    SELECT [country] AS CC, [Units Sold]
    FROM [Sales$] AS sales
    WHERE [sales.segment] = 'Government'
    ORDER BY 1
```

## JOIN clause

There are different types of joins available in Excel SQL: `INNER JOIN`, `LEFT JOIN`, `RIGHT JOIN`, `CROSS JOIN`, and `self joins`.

- `INNER JOIN` returns rows when there is a match on both tables.
- `LEFT JOIN` returns all rows from the left table even if there are no matches in the right table.
- `RIGHT JOIN` returns all rows from the right table even if there are no matches in the left table.
- `CROSS JOIN` (or `CARTESIAN JOIN`) returns the Cartesian product of the sets of records from two or more joined tables.
- `Self join` is used to join a table to itself as if the table were two tables, temporarily renaming at least one table in the SQL statement.
  You can find the syntax for the different joins below.

### INNER JOIN clause

The `INNER JOIN` keyword selects all rows from both tables as long as there is a match between the columns in both tables.

The `INNER JOIN` syntax follows this syntx pattern:

_Syntax:_

```sql
    SELECT columns
    FROM worksheet1 AS worksheet1_alias
    INNER JOIN worksheet2 AS worsksheet2_alias
    ON worksheet_alias1.column1 = worksheet2_alias.column2
```

The valid operator for the `ON` clause are `AND`, `OR`, `=`, `>`, `<`, `<>`, `>=`, `<=`, `!=`. In addition, nested joins are supported. Nested joins of more than two tables must be enclosed in parenthesis.

_Example:_

```sql
    SELECT
        A.[Segment], A.[Country], A.[Product],
        B.[Units Sold], B.[SKU]
    FROM [Products$] AS A
    INNER JOIN [Sales$] AS B
    ON A.[SKU] = B.[SKU]
```

### LEFT JOIN clause

The `LEFT JOIN` clause returns all rows from the left table (worksheet1) with the matching rows in the right table (worksheet2). The result set returns no values in the right table when there is no match.

The `LEFT JOIN` syntax follows this pattern:

_Syntax:_

```sql
    SELECT columns
    FROM worksheet1 AS worksheet1_alias
    LEFT JOIN worksheet2 AS worsksheet2_alias
    ON worksheet_alias1.column1 = worksheet2_alias.column2
```

The valid operator for the `ON` clause are `AND`, `OR`, `=`, `>`, `<`, `<>`, `>=`, `<=`, `!=`. In addition, nested joins are supported. Nested joins of more than two tables must be enclosed in parenthesis.

_Example:_

```sql
    SELECT
        A.[Segment], A.[Country], A.[Product],
        B.[Units Sold], B.[SKU]
    FROM [Products$] AS A
    LEFT JOIN [Sales$] AS B
    ON A.[SKU] = B.[SKU]
```

### RIGHT JOIN clause

The `RIGHT JOIN` clause returns all rows from the right table (worksheet2) with the matching rows in the left table (worksheet1). The result set returns no values in the left table when there is no match.

The `RIGHT JOIN` syntax follows this syntax pattern:

_Syntax:_

```sql
    SELECT columns
    FROM worksheet1 AS worksheet1_alias
    RIGHT JOIN worksheet2 AS worsksheet2_alias
    ON worksheet_alias1.column1 = worksheet2_alias.column2
```

The valid operator for the `ON` clause are `AND`, `OR`, `=`, `>`, `<`, `<>`, `>=`, `<=`, `!=`. In addition, nested joins are supported. Nested joins of more than two tables must be enclosed in parenthesis.

_Example:_

```sql
    SELECT
        A.[Segment], A.[Country], A.[Product],
        B.[Units Sold], B.[SKU]
    FROM [Products$] AS A
    RIGHT JOIN [Sales$] AS B
    ON A.[SKU] = B.[SKU]
```

### CROSS JOIN clause

The `CROSS JOIN` (or `CARTESIAN JOIN`) returns the Cartesian product of the sets of records from two or more joined tables. It equates to an inner join where the join-condition always evaluates to either True or where the join-condition is absent from the statement. Each row in the first table is paired with all the rows in the second table. This happens when there is no relationship defined between the two tables.

The `CROSS JOIN` syntax follows this pattern:

_Syntax:_

```sql
    SELECT column1, column2 FROM [worksheet1$] CROSS JOIN [worksheet2$]
```

_Example:_

```sql
    SELECT A.[E_id], B.[P_id], A.[fname]
    FROM [employee$] AS A
    CROSS JOIN [project$] AS B
```

### Self Join

`Self join` is used to join a table to itself as if the table were two tables, temporarily renaming at least one table in the SQL statement. To join a table itself means that each row of the table is combined with itself and with every other row of the table.

The Self Join syntax follows this pattern:

_Syntax:_

```sql
    SELECT column1, column2
    FROM [worksheet1$] AS alias1, [worksheet1$] AS alias2
    WHERE alias1.column1 = alias2.column2
```

_Example:_

```sql
    SELECT A.[E_id], B.[E_id]
    FROM [employee$] AS A, [employee$] AS B
    WHERE A.[E_id] = B.[Mgr_id]
```

## String Functions

The following table summarizes the most commonly used String functions used in Microsoft Excel SQL statements.

| Function | Description                                                                                                                                                                                                            |
| :------: | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `CHR` | 	Returns the character based on the ASCII value. The syntax is <br><br>`CHR( ascii_value )`<br><br>where `ascii_value` is the decimal value used to retrieve the character from the [ASCII Table](https://www.techonthenet.com/ascii/chart.php). <br><br>For example, `CHR(37)` would return `%` (i.e. the Percent Sign character).|
| `INSTR`| Returns the position of the first occurrence of a substring in a string. The syntax for the `INSTR` function in Microsoft Excel SQL is:<br><br>`INSTR( string, substring )`<br><br>where `string` is the string to search within, and `substring` is the substring which you want to find.<br><br>For example, the SQL uery <br><br>`SELECT [Segment], INSTR([Segment],'er') AS [erPosition] FROM [sheet1$]`<br><br>returns the position of the letters 'er' in the value of the `Segment` column. If `Segment` were to contain the value 'Government', the function will returns a value of 4. |
| `LCASE` | 	Converts a string to all lowercase. The syntax for the `LCASE` function in Microsoft Excel SQL is:<br><br>`LCASE( text )`<br><br>where `text` is the string which you wish to convert to lower-case. |
| `LEFT` | Extract a substring from a string, starting from the left-most character. The syntax for the `LEFT` function in Microsoft Excel SQL is:<br><br>`LEFT( text, number_of_characters )` <br><br>where `text` is the string which you wish to extract from, and `number_of_characters`  indicates the number of characters that you wish to extract starting from the left-most character. |
| `LEN` | Returns the length of the specified string. The syntax for the `LEN` function in Microsoft Excel SQL is:<br><br>`LEN( text )`<br><br>where `text` is the string that you wish to determine the length of. |
| `LTRIM` | Removes leading spaces from a string. The syntax for the `LTRIM` function in Microsoft Excel SQL is:<br><br>`LTRIM( text )`<br><br>where `text` is the string that you wish to remove leading spaces from. |
| `MID` | Extracts a substring from a string (starting at any position). The syntax for the `MID` function in Microsoft Excel SQL is:<br><br>`MID( text, start_position, number_of_characters )`<br><br>where `text` is the string which you wish to extract from, `start_position`  indicates the position in the string that you will begin extracting from (the first position in the string is 1), and `number_of_characters` is the number of characters that you wish to extract. |
| `REPLACE` | Replaces a sequence of characters in a string with another set of characters. The syntax for the REPLACE function in Microsoft Excel SQL is:<br><br>`REPLACE ( string1, find, replacement, [start, [count]] )`<br><br>Where `string1` is the string to replace a sequence of characters with another set of characters, `find` is the string that will be searched for in `string1`, `replacement` is the string which will replace `find` in `string1`. <br><br>For example, the SQL statement `SELECT REPLACE([Segment],' ','_') FROM [sheet1$]` would change all space characters in the `Segment` column values to underscores (a value such as 'Channel Partners' would be returned as 'Channel_Partners').<br><br>Parameter `start` is optional, and specifies the position in `string1` to begin the search. If the `start` parameter is omitted, the `REPLACE` function will begin the search at position 1. Parameter `count` is also optional, and specifies the the number of occurrences to replace. If the `count` parameter is omitted, the `REPLACE` function will replace all occurrences of `find` with `replacement`. If you wish to specify the `count` parameter, you must also specify the `start` parameter.<br><br>For example, the SQL statement `SELECT REPLACE([Segment],'e','X',1,2)FROM [sheet1$]` would change the first 2 occurences of the letter 'e' in the `Segment` values to the letter 'X' (A value such as 'Government' becomes 'GovXrnmXnt', while a value of 'Enterprise' becomes 'XntXrprise' )|
| `RIGHT` | Extract a substring from a string, starting from the right-most character. The syntax for the `RIGHT` function in Microsoft Excel SQL is:<br><br>`RIGHT( text, number_of_characters )`<br><br>where `text` is the string which you wish to extract from, and `number_of_characters`  indicates the number of characters that you wish to extract starting from the right-most character. |
| `RTRIM` | Removes trailing spaces from a string. The syntax for the `RTRIM` function in Microsoft Excel SQL is:<br><br>`RTRIM( text )`<br><br>where `text` is the string that you wish to remove trailing spaces from. |
| `SPACE` | Returns a string value with a specified number of spaces. The syntax for the `SPACE` function in Microsoft Excel SQL is:<br><br>`SPACE( number )`<br><br>where `number` is the number of spaces to be returned. |
| `STR` | Returns a string representation of a number. The syntax for the `STR` function in Microsoft Excel SQL is:<br><br>`STR( number )`<br><br>where `number` is the numeric value that you wish to convert to a string.|
| `TRIM` | Returns a text value with the leading and trailing spaces removed. The syntax for the `TRIM` function in Microsoft Excel SQL is:<br><br>`TRIM( text )`<br><br>where `text` is the string that you wish to remove leading and trailing spaces from. |
| `UCASE` | Converts a string to all uppercase. The syntax for the `UCASE` function in Microsoft Excel SQL is:<br><br>`UCASE( text )`<br><br>where `text` is the string which you wish to convert to upper-case. |

## Resources/Acknowlegements

The information in this document was derived from the following sources:

- [Writing SQL Queries against Excel files](https://querysurge.zendesk.com/hc/en-us/articles/205766136-Writing-SQL-Queries-against-Excel-files)
- [MS Excel SQL Query Reference](https://www.aquaclusters.com/app/home/project/public/aquadatastudio/wikibook/MS-Excel/page/SQL-Query-Reference/SQL-Query-Reference)
- [Fundamental Microsoft Jet SQL for Access 2000](https://docs.microsoft.com/en-us/previous-versions/office/developer/office2000/aa140011%28v%3doffice.10%29)
- [Intermediate Microsoft Jet SQL for Access 2000](https://docs.microsoft.com/en-us/previous-versions/office/developer/office2000/aa140015%28v%3doffice.10%29)
- [The Power of SQL Applied to Excel Tables for Fast Results](https://morsagmon.com/blog/The-Power-of-SQL-Applied-to-Excel-Tables-for-Fast-Results/)
- [Format function](https://docs.microsoft.com/en-us/office/vba/Language/Reference/User-Interface-Help/format-function-visual-basic-for-applications)
- [VBA-(SQL)-String Video-Tutorial for Beginner](https://codekabinett.com/rdumps.php?Lang=2&targetDoc=vba-sql-string-tutorial)
- [SQL Tutorial for Beginners](https://database.guide/sql-tutorial-for-beginners/)
- [Null Values](https://bettersolutions.com/vba/sql/null-values.htm)
- [MS Excel: Formulas and Functions](https://www.techonthenet.com/excel/formulas/)

Microsoft provives a [free Excel spreadsheet containing sample data](https://go.microsoft.com/fwlink/?LinkID=521962) which provides a useful source of Excel data for writing practice queries.

---

<center>

Like this tool? [Buy me a coffee! ☕](https://www.buymeacoffee.com/exceltographviz)

</center>

