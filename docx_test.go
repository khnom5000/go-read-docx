package docx

import (
	"strings"
	"testing"
)

const testFile = "./TestDocument.docx"

func getDocument() (d Document) {
	d, err := GetDocument(testFile)
	if err != nil {
		panic(err)
	}
	return d
}

func TestGetDocuemnt(t *testing.T) {
	_, err := GetDocument(testFile)
	if err != nil {
		t.Error("Failed with error: ", err)
	}
}

func TestParagraphs(t *testing.T) {
	d := getDocument()
	p := strings.Join(d.Body.Paragraphs, " ")
	if !strings.Contains(p, "Start of page one") {
		t.Error("Failed to find 'Start of page one'")
	}
}

func TestTable(t *testing.T) {
	d := getDocument()

	tbl := d.Body.Tables[0].TableRows
	var table [][]string
	for _, r := range tbl {
		var row []string
		for _, c := range r.TableColumns {
			row = append(row, c.Cell)
		}
		table = append(table, row)
	}

	row := strings.Join(table[0], " ")
	if !strings.Contains(row, "1 2 3 4") {
		t.Error("Failed to find '1 2 3 4' in Table[0]")
	}
}

func TestHeader(t *testing.T) {
	h, err := GetHeader(testFile)
	if err != nil {
		t.Error("Failed with error: ", err)
	}

	if !strings.Contains(h.Text, "This is a header") {
		t.Error("Failed to find 'This is a header' in header")
	}
}

func TestFooter(t *testing.T) {
	f, err := GetFooter(testFile)
	if err != nil {
		t.Error("Failed with error: ", err)
	}

	if !strings.Contains(f.Text, "This is a footer") {
		t.Error("Failed to find 'This is a footer' in footer")
	}
}
