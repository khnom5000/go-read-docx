package docx

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"strings"
)

//Body : Elements found within the body section of a word docx
type Body struct {
	Paragraph []string `xml:"p>r>t"`
	Table     Table    `xml:"tbl"`
}

//Table : Simply holds the tags that tbl can hold
type Table struct {
	TableRow []TableRow `xml:"tr"`
}

//TableRow : Holds all the things a tr element can
type TableRow struct {
	TableColumn []TableColumn `xml:"tc"`
}

//TableColumn : Holds all the things a tc element can
type TableColumn struct {
	Cell []string `xml:"p>r>t"`
}

//Document : Holds all the things a docx Document can
type Document struct {
	XMLName xml.Name `xml:"document"`
	Body    Body     `xml:"body"`
}

//Extract : parse the xml into the structures defined as per "Document"
func Extract(xmlContent string) (d Document) {
	err := xml.Unmarshal([]byte(xmlContent), &d)
	if err != nil {
		panic(err)
	}
	return
}

//UnpackDocx : Unzip the doc file
func UnpackDocx(filePath string) (*zip.ReadCloser, []*zip.File) {
	reader, err := zip.OpenReader(filePath)
	if err != nil {
		panic(err)
	}
	return reader, reader.File
}

//WordDocToString : This converts the file interface object into a raw string
func WordDocToString(reader io.Reader) (content string) {
	_content := make([]string, 100)
	data := make([]byte, 100)

	for {
		n, err := reader.Read(data)
		_content = append(_content, string(data))
		if err == io.EOF && n == 0 {
			break
		}
	}
	content = strings.Join(_content, "")
	return
}

//RetrieveWordDoc : Simply loops over the files looking for the file with name "word/document"
func RetrieveWordDoc(files []*zip.File) (file *zip.File) {
	for _, f := range files {
		if f.Name == "word/document.xml" {
			file = f
		}
	}
	return
}

//OpenWordDoc : Opens the zipfile into a reader for WordDocToString
func OpenWordDoc(doc zip.File) (rc io.ReadCloser) {
	rc, err := doc.Open()
	if err != nil {
		panic(err)
	}
	return
}
