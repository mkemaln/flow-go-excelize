# Fllow-Go-Excelize

This is a Go API Based flowchart maker in excel file output using Go Excelize. I mainly develop this for another project but i will publicly share here.

Right now the prototype is still using URL query based input which there are:
|Query|Example|Usage|
|--|--|--|
|start*|G6|starting cell of the flowchart|
|width*|int|width of the shape in all general|
|height*|int|height of the shape in all general|
|[gap*](##gap)|int|how much row gap for each shape|
|pad*|int|padding for the cell on the each shape |
|[orders*](##order)|[1,2,2,4]|order of the shape, those number is depicting the column position|
|[shapes*](##shapes)|[rect,ellipse]|name of the shape based on the go excelize docs|

note: asterisk or * is a required query

## How to run?

you can compile it or simply run the main.go file in the [cmd/server/main.go](cmd/server/main.go)

```
go run cmd/server/main.go
```

then input each query the 
```
http://localhost:8080/excel?start=&width=&height=&gap=&pad=&orders=&shapes=
```

the output file is a random name excel file.

# Changelog / Update

#### v0.0.4 - 10/11/2025
- **fixed**
- change the placement of upperLeftConn and upperRightConn so it wont stack with other line
- make only the decision diagram that will printed the `true` line from below
- fixing the decision false connector one if the target in the same column
- fixed the width, height, gap input to 120, 65, 30
- **planned TODO:**
- fixing line that were bad rendered in multiple condition (different gap)
- experimenting more

example non stacked flow: [result](https://files.catbox.moe/1uuqcu.png)

```
http://localhost:8080/excel?shapes=rect,flowChartDecision,rect,rect,flowChartDecision,rect&start=G6&orders=1,4,3,4,2,3&width=120&height=65&pad=30&gap=1&false_branches=1:0,4:3&true_branches=1:2,4:5
```

example stacked flow: [result](https://files.catbox.moe/szm1x2.png)

```
http://localhost:8080/excel?shapes=rect,flowChartDecision,rect,rect,flowChartDecision,rect&start=G6&orders=1,1,1,1,1,1&width=120&height=65&pad=30&gap=1&false_branches=1:0,4:3&true_branches=1:2,4:5
```

#### v0.0.3 - 05/11/2025
- fixing the calculation of the line placing for upperLeftConn and upperRightConn
- **planned TODO:**
- change the placement of upperLeftConn and upperRightConn so it wont stack with other line
- make only the decision diagram that will printed the `true` line from below
- fixing the decision false connector one if the target in the same column
- experimenting more

example: [result](https://files.catbox.moe/iq32jg.png)

```
http://localhost:8080/excel?shapes=rect,flowChartDecision,rect,rect,flowChartDecision,rect&start=G6&orders=1,3,2,3,2,4&width=80&height=40&pad=10&gap=1&false_branches=1:0,4:3&true_branches=1:2,4:5
```

#### v0.0.2 - 03/11/2025
- change the mechanism of the decision diagram connection
- connection method for decision diagram is using index-to-index based mapping, i.e 1:0 => means that index 1 shape (first decision) is connected to the index 0 (first rect shape) and will use that connection to make a connection line 
- increase the lineWidth so the tip arrow of arrow shape is visible
- **ISSUE rn:**
- still trying the workaroud so the line is printed first the the shape to prevent the stacking line over a shape
- the placement logic of connecting to the upper shape is still at development, i hope for that part is the connection will consist of 3 lines like
```

     <-|
       |
    ---|

```

example: [result](https://files.catbox.moe/0euncq.png)

```

http://localhost:8080/excel?shapes=rect,flowChartDecision,rect,rect,flowChartDecision,rect&start=G6&orders=1,2,3,3,4,4&width=80&height=40&pad=10&gap=1&false_branches=1:0,4:2&true_branches=1:2,4:5
```

#### v0.0.1 - 17/09/2025
- Implemented shape rendering with down order
- Implemented URL query based input for making an excel file
- example query : [result](https://files.catbox.moe/pxsl9m.png)

```
http://localhost:8080/excel?width=80&height=40&shapes=rect,rect,rect,flowChartDecision&start=D4&gap=1&pad=10&orders=1,1,2,4
```

# Additional Explanation
## gap
gap explaination here

## orders
orders explaination here

## shapes
shapes explaination here