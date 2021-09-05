package docx

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"io"
)

//Body : Elements found within the body section of a word docx
type Body struct {
	Paragraphs []string `xml:"p>r>t"`
	Tables     []Table  `xml:"tbl"`
}

//Table : Simply holds the tags that tbl can hold
type Table struct {
	TableRows []TableRow `xml:"tr"`
}

//TableRow : Holds all the things a tr element can
type TableRow struct {
	TableColumns []TableColumn `xml:"tc"`
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

//UnpackDocx : Unzip the doc file
func UnzipDocx(f string) (*zip.ReadCloser, error) {
	rc, err := zip.OpenReader(f)
	if err != nil {
		return nil, err
	}
	return rc, nil
}

//GetWordDoc : loops over the files looking for the file with name "word/document.xml"
//Can only find 1 "word/document.xml" if 0 or more than 1 are found an err will be thrown
func GetWordDoc(files *zip.ReadCloser) (*zip.File, error) {
	c := 0
	f := new(zip.File)
	for _, file := range files.File {
		if file.Name == "word/document.xml" {
			f = file
			c++
		}
	}

	if c == 0 {
		return nil, errors.New("no \"word/document.xml\" found in file")
	} else if c > 1 {
		return nil, errors.New("more than one \"word/document.xml\" found in file")
	}

	return f, nil
}

//OpenWordDoc : Opens the zipfile into a reader for WordDocToString
func OpenWordDoc(doc zip.File) (io.ReadCloser, error) {
	rc, err := doc.Open()
	if err != nil {
		return nil, err
	}
	return rc, nil
}

//WordDocToString : This converts the file interface object into a raw string
func WordDocToString(reader io.Reader) (string, error) {
	content, err := io.ReadAll(reader)
	if err != nil {
		return "", err
	}
	return string(content), nil
}

//Extract : parse the xml into the structures defined as per "Document"
func Extract(xmlContent string) (Document, error) {
	d := Document{}
	err := xml.Unmarshal([]byte(xmlContent), &d)
	if err != nil {
		return Document{}, err
	}
	return d, nil
}

//GetDocument : Wrapper function that takes a file, and retuns both the Document{} and ReadCloser
func GetDocument(f string) (Document, error) {
	reader, err := UnzipDocx(f)
	if err != nil {
		return Document{}, err
	}
	defer reader.Close()

	docFile, err := GetWordDoc(reader)
	if err != nil {
		return Document{}, err
	}

	doc, err := OpenWordDoc(*docFile)
	if err != nil {
		return Document{}, err
	}
	defer doc.Close()

	content, err := WordDocToString(doc)
	if err != nil {
		return Document{}, err
	}

	d, err := Extract(content)
	if err != nil {
		return Document{}, err
	}
	return d, nil
}
