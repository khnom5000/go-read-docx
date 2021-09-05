package main

import (
	"fmt"

	docx "github.com/khnom5000/go-read-docx"
)

func main() {

	d, err := docx.GetDocument("./examples/TestDocument.docx")
	if err != nil {
		panic(err)
	}

	fmt.Println("Show all paragraphs")
	ps := d.Body.Paragraphs
	for i, p := range ps {
		fmt.Println("Para:", i, p)
	}

	fmt.Println("Show first table")
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

	fmt.Println("Show all tables")
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

	fmt.Println("Show Header")
	h, err := docx.GetHeader("./examples/TestDocument.docx")
	if err != nil {
		panic(err)
	}
	fmt.Println(h.Text)

	fmt.Println("Show Footer")
	f, err := docx.GetFooter("./examples/TestDocument.docx")
	if err != nil {
		panic(err)
	}
	fmt.Println(f.Text)

}
