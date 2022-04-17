# go-read-docx
## simple way to read a docx in golang, including support for tables

```go
package main

import (
	docx "github.com/khnom5000/go-read-docx"
)

func main() {
	d, reader, err := docx.GetDocument("./TestDocument.docx")
	if err != nil {
		panic(err)
	}
}
```

### Example - Read text in a docx:

Code:
```go
...
ps := d.Body.Paragraphs
for i, p := range ps {
	fmt.Println("Para:", i, p)
}
...
```

### Example - Read a table thats inside a docx:

Input:
```
+-----+---+---+----+
|   1 | 2 | 3 |  4 |
|   8 | 8 | 8 | 66 |
| 123 | 1 | 1 |  1 |
|     |   |   |    |
+-----+---+---+----+
```
Code:
```go
...
t := d.Body.Tables[0].TableRows
var table [][]string
for _, r := range t {
	var row []string
	for _, c := range r.TableColumns {
		row = append(row, c.Cell)
	}
	table = append(table, row)
}
fmt.Println(table)
...
```

Output:
```
[[1 2 3 4] [8 8 8 66] [123 1 1 1] [   ]]
```

### Example - Read more than one table in the same docx:

Input:
```
+-----+-----+-----+-----+
|   1 |   2 |   3 |   4 |
|   8 |   8 |   8 |  66 |
| 123 |   1 |   1 |   1 |
|     |     |     |     |
+-----+-----+-----+-----+
...
+-----+-----+-----+-----+
|   7 |   8 |   9 |   0 |
|   0 |  33 |  66 |  99 |
| 123 | 100 | 100 | 100 |
|     |     |     |     |
+-----+-----+-----+-----+
```
Code:
```go
...
ts := d.Body.Tables
for _, t := range ts {
	var table [][]string
	for _, tr := range t.TableRows {
		var row []string
		for _, tc := range tr.TableColumns {
			row = append(row, tc.Cell)
		}
		table = append(table, row)
	}
	fmt.Println(table)
}
...
```
Output:
```
[[1 2 3 4] [8 8 8 66] [123 1 1 1] [   ]]
[[7 8 9 0] [0 33 66 99] [123 100 100 100] [   ]]
```

### Example - Get the Headers:

```go
...
h, err := docx.GetHeader("./TestDocument.docx")
if err != nil {
	panic(err)
}
fmt.Println(h.Text)
...
```
The above also works for footers just swap GetHeader() with GetFooter()

Output:
```
Show all paragraphs
Para: 0 Start of page one.
Para: 1 This is a table.
Para: 2 This is the second table.
Para: 3 This is the end of the doc!
Show first table
[[1 2 3 4] [8 8 8 66] [123 1 1 1] [   ]]
Show all tables
[[1 2 3 4] [8 8 8 66] [123 1 1 1] [   ]]
[[7 8 9 0] [0 33 66 99] [123 100 100 100] [   ]]
Show Header
This is a header.
Show Footer
This is a footer.
```

[![Go Reference](https://pkg.go.dev/badge/github.com/khnom5000/go-read-docx.svg)](https://pkg.go.dev/github.com/khnom5000/go-read-docx)