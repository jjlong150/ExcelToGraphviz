# Constructing Simple Graphs

## Connect 3 Nodes

Next, lets expand upon the graph we just created to have additional relationships. Assume that:

-   'a' is related to 'b' (already drawn)
-   'b' is related to 'c'
-   'c' is related to 'a'

The Excel data appears as shown on rows 3-5. Press the ```Refresh Graph``` button, and the Excel worksheet now looks like:

![](../media/f4f912714826d55f8e73d9b767f4a088.png)


*Graphviz Source*

``` dot
    strict digraph "main" 
    {
        layout="dot";
        rankdir="TB";

        "a" -\> "b";
        "b" -\> "c";
        "c" -\> "a"; 
    } 
```

## Add Edge Labels

Now, let us add data into the ```Label``` column to label the relationships. Fill in Column D as shown below. Press the ```Refresh Graph``` button, and the Excel worksheet now looks like:

![](../media/8f0481849c081a24edc4a502224161e5.png)

*Graphviz Source*

```dot
    strict digraph "main" 
    { 
        layout="dot"; 
        rankdir="TB";

        "a" -\> "b"[ label="is related to" ]; 
        "b" -\> "c"[ label="is related to" ]; 
        "c" -\> "a"[ label="is related to" ]; 
    }
```

## Add Node Labels

The graph is how we want to see it, but the nodes need to be labeled. We do not want to change all our edges; however, we would like to replace 'a' with 'Alpha', 'b' with 'Bravo', and 'c' with 'Charlie'. The Relationship Visualizer assumes that when there is information in the ```Item``` column, but not in the ```Related Item``` column that the data corresponds to a node.

To label the nodes we will add 3 node definitions to the "data worksheet (rows 6, 7, 8) and press the ```Refresh Graph``` button. The Excel worksheet now looks like:

![](../media/3bd5c434221f90b9ea8c636eda70ccf3.png)

*Graphviz Source*

```dot
    strict digraph "main" 
    { 
        layout="dot"; 
        rankdir="TB";

        "a" -\> "b"[ label="is related to" ]; 
        "b" -\> "c"[ label="is related to" ]; 
        "c" -\> "a"[ label="is related to" ]; 
        "a"[ label="Alpha" ]; 
        "b"[ label="Bravo" ]; 
        "c"[ label="Charlie" ]; 
    }
```

## Specify Ports

Graphviz decides what it thinks is the best placement of the head and tail of an edge to produce a balanced graph.

Sometimes you might want to control where the edges begin or end. You can do that by specifying a port on the ```Item``` or ```Related Item``` ID, in the same manner as a URL. Ports are identified by a colon character ```:``` and then a compass point ```n```, ```s```, ```e```, ```w```, ```ne```, ```nw```, ```se```, ```sw``` or ```c``` for center.

If we change row 5 from the example above to have the edge from "c" to "a" exit from the south port of "c", the ```Item``` is now specified as ```c:s```, and the Excel data is changed slightly as shown in row 5. Press the ```Refresh Graph``` button, and the Excel worksheet now looks like:

![](../media/87a90b140ec0d5987a284daa8abd19cf.png)


*Graphviz Source*

```dot
    strict digraph "main"
    {
        layout="dot";
        rankdir="TB";
        
        "a" -> "b"[ label="is related to" ];
        "b" -> "c"[ label="is related to" ];
        "c":"s" -> "a"[ label="is related to" ];
        "a"[ label="Alpha" ];
        "b"[ label="Bravo" ];
        "c"[ label="Charlie" ];
    }
```

## Specify Clusters

If you wish to cluster some elements of the graph you can do so by adding a row with an open brace "{" in the ```Item``` column above the first row of data to be placed in the group and provide a title for the cluster in the ```Label``` column. Next, add row with a close brace "}" in the ```Item``` column after the last row of data.

For example, this Excel worksheet does not have clusters.

![](../media/d0011b67a73a9e14312423b01c73fcfb.png)

*Graphviz Source*

```dot
    strict digraph "main"
    {
        layout="dot";
        rankdir="TB";
        
        "start" -> "a0";
        "a0" -> "a1";
        "a1" -> "a2";
        "a2" -> "end";
    }
```

To cluster nodes a0, a1, and a2, calling the cluster "process \#1" the worksheet is revised to add an open brace {with the label "process \#1" on row 3, and a close brace } on rows 6 as follows.

Press the ```Refresh Graph``` button, and the Excel worksheet now looks like:

![](../media/7f02cd43f77aa9e1cd511d5e443b3bdf.png)

*Graphviz Source*

```dot
    strict digraph "main"
    {
        layout="dot";
        rankdir="TB";
        
        "start" -> "a0";
        subgraph "cluster_1" {  label="process #1"
            "a0" -> "a1";
            "a1" -> "a2";
        }
        "a2" -> "end";
    }
```
## Specify Clusters Within Clusters

Graphviz permits clusters within clusters. Let us extend the example by adding an additional set of braces to cluster the relationship between a1 and a2. We will insert a new row 5 placing an open brace { in the ```Item``` column with the Label column set to "process \#2", and a new row 7 with a close brace } in the ```Item``` column.

Press the ```Refresh Graph``` button, and the Excel worksheet now looks like:

![](../media/1df108aa9f36e24f4f7958f5fe999189.png)

*Graphviz Source*

```dot
    strict digraph "main"
    {
        layout="dot";
        rankdir="TB";
        
        "start" -> "a0";
        subgraph "cluster_1" {  label="process #1"
            "a0" -> "a1";
            subgraph "cluster_2" {  label="process #2"
                "a1" -> "a2";
            }
        }
        "a2" -> "end";
    }
```

Graphviz does not limit the number of clusters you can have. In this example, we have added rows 10-14 to insert an additional cluster labeled "process \#3".

Press the ```Refresh Graph``` button, and the Excel worksheet now looks like:

![](../media/0edd4afd935217ae92566ab83893fae8.png)

*Graphviz Source*

```dot
    strict digraph "main"
    {
        layout="dot";
        rankdir="TB";
        
        "start" -> "a0";
        subgraph "cluster_1" {  label="process #1"
            "a0" -> "a1";
            subgraph "cluster_2" {  label="process #2"
                "a1" -> "a2";
            }
        }
        "a2" -> "end";
        "start" -> "b0";
        subgraph "cluster_3" {  label="process #3"
            "b0" -> "b1";
        }
        "b1" -> "end";
    }
```

What is important to note is that you must ensure that you have an equal number of open braces as you do close braces. **If you have a mismatch between the number of open and close braces, then Graphviz will not draw the graph.**

## Specify Comma-separated Items

Another feature of the Relationship Visualizer is the ability to specify a comma-separated list of Item names and have a relationship created for each Item. For example, we can say that Mr. Brady is the father of Greg, Peter, and Bobby on one row as follows:

![](../media/d58e637f465efc9ac6a115a7077d477a.png)

*Graphviz Source*

```dot
    strict digraph "main"
    {
        layout="dot";
        rankdir="TB";
        
        "Mr. Brady" -> "Greg"[ label="Father of" ];
        "Mr. Brady" -> "Peter"[ label="Father of" ];
        "Mr. Brady" -> "Bobby"[ label="Father of" ];
    }
```

The comma-separated list can also appear in the ```Item``` column, such as:

![](../media/220ca8476484163f0a3de41b90ad84be.png)

*Graphviz Source*

```dot
    strict digraph "main"
    {
        layout="dot";
        rankdir="TB";
        
        "Marcia" -> "Mrs. Brady"[ label="Daughter of" ];
        "Jan" -> "Mrs. Brady"[ label="Daughter of" ];
        "Cindy" -> "Mrs. Brady"[ label="Daughter of" ];
    }
```


Or a comma-separated list can be used in both the ```Item```, and the ```Related Item``` column such as the parental relationship below:

![](../media/ac01a7b46880bb75a0764b30bbbf38bb.png)

*Graphviz Source*

```dot
    strict digraph "main"
    {
        layout="dot";
        rankdir="TB";
        
        "Mr. Brady" -> "Greg";
        "Mr. Brady" -> "Peter";
        "Mr. Brady" -> "Bobby";
        "Mr. Brady" -> "Marcia";
        "Mr. Brady" -> "Jan";
        "Mr. Brady" -> "Cindy";
        "Mrs. Brady" -> "Greg";
        "Mrs. Brady" -> "Peter";
        "Mrs. Brady" -> "Bobby";
        "Mrs. Brady" -> "Marcia";
        "Mrs. Brady" -> "Jan";
        "Mrs. Brady" -> "Cindy";
    }
```
