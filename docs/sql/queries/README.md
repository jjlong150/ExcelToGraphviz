---
prev: /sql/
next: /sql/extensions/
---

# SQL Queries for Creating Graphs

## Introduction
In the Relationship Visualizer download package, you’ll find a directory named `15 - Using SQL - World Map` inside the `samples` folder. We’ll use this example to demonstrate how to work with the `sql` worksheet.

The `SQL - World Map` directory contains:

- A copy of the Relationship Visualizer workbook with SQL statements already defined  
- Preconfigured styles on the `styles` worksheet  
- A `countries.xlsx` workbook containing the data we will query  

An example of the `countries` workbook appears as follows:

| ![](./countries.png) |
| -------------------- |

## Our Use Case

Our goal is to graph the countries connecting each country to the continent, and using unique colors for each continent. 

## Style Elements

If we file on the `Continent` column we see there are 7 continents.

| ![](./continent-filter.png) |
| -------------------- |

Using the [Style Designer](../../designer/) we create a unique style for each country by assigning different Fill Colors. The styles are stored on the [styles](../../styles/) worksheet using the Continent Name as the Style Name.

This snipped from the `styles` worksheet shows the 7 node style definitions which correspond to the continent names.

| ![](./continent-styles.png) |
| -------------------- |

## Building the SQL Queries and Directives

Next, open `Relationship Visualizer.xlsm` workbook. From the Launchpad ribbon tab select `SQL`. 

| ![](./launchpad-sql.png) |
| -------------------- |

The `sql` worksheet will appear. You’ll rows containing SQL statements, and some additional directives which. These queries extract values from the `countries` workbook and are used to generate a graph showing continents, countries, and their shared borders.

The `sql` worksheet appears as follows:

![](./sql-worksheet.png)

### Specify the Data Source

The first directive statement on row 4 is `SET DATA FILE`. It informs the Relationship Visualizer to use the file specified in column C (`countries.xlsx`) to be the default data source.

### Create the Country Nodes

The first SQL statement on row 7 selects the column `ISO` to represent the `Item`, as well as the node `label`. The `Country` name is selected to be an `External Label`, and the `Continent` name is selected to be the `Style Name`. 

The SQL is written as follows:

```sql
SELECT [ISO]       as [Item],
       [ISO]       as [Label],
       [Country]   as [External Label],
       [Continent] as [Style Name]
FROM  [countries$]
WHERE [ISO] IS NOT NULL
ORDER BY [Country] ASC
```

### Create the Continent Nodes

The second SQL statement on row 10 will extract the 7 seven unique continent names from the list of countries by using the `DISTINCT` clause. The `Continent` name will be used as the `Item ID` as well as the node `label`. A `Style Name` of **Continent** will be used for all the rows.

The SQL is written as follows:

```sql
SELECT DISTINCT [Continent] as [Item],
                [Continent] as [Label],
                'Continent' as [Style Name]
FROM [countries$]
ORDER BY [Continent]
```

### Create the Country-to-Country Edges

The third SQL statement on row 13 creates edge relationship rows by selecting the ISO value as the `Item ID`, and the `Neighbors` column as the `Related Item` value. The `Neighbors` column in the source worksheet contains comma-delimited `ISO` values which are neighboring countries to that row. The Relationship Visualizer has built-in logic to expand the comma-delimited list into multiple relationships. No style name is provided, so the default edge style will be used. Note that on the `WHERE` clause the directive `IS NOT NULL` has been added. This clause causes the query to skip rows with empty cells, as there are no values with which to express relationships.

The SQL is written as follows:

```sql
SELECT [ISO]       as [Item],
       [Neighbors] as [Related Item]
FROM   [countries$]
WHERE [Neighbors] IS NOT NULL
```

### Create the Continent-to-Country Edges

The fourth SQL statement on row 16 is used to group countries by continent. It creates edge relationships by placing the Continent name as the `Item ID`, and the ISO country code in the `Related Item` column. The `Style Name` is specified as **ContinentToCountry** which has been defined as an invisible edge.

The SQL is written as follows:

```sql
SELECT [Continent]          as [Item],
       [ISO]                as [Related Item],
       'ContinentToCountry' as [Style Name]
FROM [countries$]
WHERE [ISO] IS NOT NULL
```

### Run the SQL Queries

At this point our queries are complete. Press the `Run SQL Commands` button.

| ![](./run-sql-button.png) |
| ------------------------- |

The SQL commands are run in sequence from top to bottom. Results are written to the `data` worksheet, and the query result status is displayed in column D such as:

| ![](./query-status.png) |
| ----------------------- |

If we switch worksheets to the `data` worksheet, it appears as follows:

| ![](./data-results.png) |
| ----------------------- |

The data is all present and in the appropriate columns for graphing. In this example **678 rows of data have been created using 4 SQL statements!** 

### Graph the Data

Press the `Refresh` button to graph the data. Since this is a large data set, be prepared to wait a little while for the results.

When Graphviz completes its work, you should see a logical world graph which appears as follows :

| ![](./data-graph.png)   |
| ----------------------- |

*Options used: Worksheet = `graph`, Zoom = 45%, layout=fdp, splines=true, graphtype=undirected.*

In a few short minutes we have gone from tabular Excel data to a graph visualization.

## Directives

The Relationship Visualizer has additional directives which look like SQL which can be used to enhance the batch processing capabilities. Think of them as extensions to SQL.

### Clear the results from the `data` worksheet

- `RESET`
  
### Specify the workbook containing the data

- `SET DATA FILE` where the workbook name is placed in Column C.
  
### Show the graph immediately after all SQL statements processed
Pattern: `PREVIEW [ AS ( DIRECTED | UNDIRECTED ) GRAPH ]` 

- `PREVIEW`
- `PREVIEW AS DIRECTED GRAPH` 
- `PREVIEW AS UNDIRECTED GRAPH` 

### Publish the graphs as files
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

---

<center>

Like this tool? [Buy me a coffee! ☕](https://www.buymeacoffee.com/exceltographviz)

</center>
