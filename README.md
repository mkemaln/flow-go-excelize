# Fllow-Go-Excelize

This is a Go API Based flowchart maker in excel file output using Go Excelize. I mainly develop this for another project but i will publicly share here.

Right now the prototype is still using URL query based input which there are:
|Query|Example|Usage|
|--|--|--|
|start|G6|starting cell of the flowchart|
|width|int|width of the shape in all general|
|height|int|height of the shape in all general|
|[gap](##gap)|int|how much row gap for each shape|
|pad|int|padding for the cell on the each shape |
|[orders](##order)|[1,2,2,4]|order of the shape, those number is depicting the column position|
|[shapes](##shapes)|[rect,ellipse]|name of the shape based on the go excelize docs|

# Changelog / Update

#### v0.0.1 - 17/09/2025
- Implemented shape rendering with down order
- Implemented URL query based input for making an excel file
- example query : [result](https://files.catbox.moe/pxsl9m.png)

```
http://localhost:8080/excel?width=80&height=40&shapes=rect,rect,rect,flowChartDecision&start=D4&gap=1&pad=10&orders=1,1,2,4
```