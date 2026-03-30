# Directives

The Relationship Visualizer has additional directives which look like SQL which can be used to enhance the batch processing capabilities. Think of them as extensions to SQL.

## Clear the results from the `data` worksheet

- `RESET`
  
## Specify the workbook containing the data

- `SET DATA FILE` where the workbook name is placed in Column C.
  
## Placeholders
Set placeholder values for substitution into SQL, and wrap placeholder name in braces in the SQL.

- `SET PLACEHOLDER placeholder_name = value`

For example, to query for records matching the value "Metropolitan", assign Metropolitan to a placeholder, 
then use the placeholder name wrapped in braces in the SQL as static text in WHERE comparisons, label, or tooltip text.

```sql
SET PLACEHOLDER subway_line = Metropolitan

SELECT [station_from]       AS [Item], 
       [station_to]         AS [Related Item], 
       '{subway_line} Line' AS [Label], 
FROM   [london_underground$]
WHERE  [tube_route] = '{subway_line}'
```

## Show the graph immediately after all SQL statements processed
Pattern: `PREVIEW [ AS ( DIRECTED | UNDIRECTED ) GRAPH ]` 

- `PREVIEW`
- `PREVIEW AS DIRECTED GRAPH` 
- `PREVIEW AS UNDIRECTED GRAPH` 

## Publish the graphs as files
Pattern: `PUBLISH [ ALL VIEWS ] [ AS ( DIRECTED | UNDIRECTED ) GRAPH ] [ file prefix ]`

- `PUBLISH`
- `PUBLISH AS DIRECTED GRAPH`
- `PUBLISH AS UNDIRECTED GRAPH`
- `PUBLISH ALL VIEWS`
- `PUBLISH ALL VIEWS AS DIRECTED GRAPH`
- `PUBLISH ALL VIEWS AS UNDIRECTED GRAPH`
  
You can specify a value to use for the File Name Prefix as the last value of the directive. for example, if the desired prefix is `foobar`, the directives are:
- `PUBLISH foobar`
- `PUBLISH AS DIRECTED GRAPH foobar`
- `PUBLISH AS UNDIRECTED GRAPH foobar`
- `PUBLISH ALL VIEWS foobar`
- `PUBLISH ALL VIEWS AS DIRECTED GRAPH foobar`
- `PUBLISH ALL VIEWS AS UNDIRECTED GRAPH foobar`

## Logging
Turn on/off logging of errors to file `Relationship Visualizer ADO Log.txt`
- `ENABLE LOGGING`
- `DISABLE LOGGING`

## Environment Diagnostics
Write environment diagonostic information to file `Relationship Visualizer ADO Log.txt`
- `LOG ENVIRONMENT`

