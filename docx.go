package docx

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
)

//Body : Elements found within the body section of a word docx
type Body struct {
	Paragraphs []string `xml:"p>r>t"`
	Tables     []Table  `xml:"tbl"`
}

//Table : Holds all the things a tbl element can
type Table struct {
	TableRows []TableRow `xml:"tr"`
}

//TableRow : Holds all the things a tr element can
type TableRow struct {
	TableColumns []TableColumn `xml:"tc"`
}

//TableColumn : Holds all the things a tc element can
type TableColumn struct {
	Cell string `xml:"p>r>t"`
}

//Document : Holds all the things a docx document(.xml) can
type Document struct {
	XMLName xml.Name `xml:"document"`
	Body    Body     `xml:"body"`
}

//Header : Holds all the things a docx header1(.xml) can
type Header struct {
	XMLName xml.Name `xml:"hdr"`
	Text    string   `xml:"p>r>t"`
}

//Footer : Holds all the things a docx footer1(.xml) can
type Footer struct {
	XMLName xml.Name `xml:"ftr"`
	Text    string   `xml:"p>r>t"`
}

var docxPathOfDocumentXML = "word/document.xml"
var docxPathOfHeaderXML = "word/header1.xml"
var docxPathOfFooterXML = "word/footer1.xml"

//unzipDocx : Unzip the doc file
func unzipDocx(f string) (*zip.ReadCloser, error) {
	rc, err := zip.OpenReader(f)
	if err != nil {
		return nil, err
	}
	return rc, nil
}

//getXMLFile : loops over the files looking for the file matching the provided xmlDoc string
//Can only find 1 occurance of the xmlDoc file if 0 or more than 1 are found an err will be thrown
func getXMLFile(files *zip.ReadCloser, xmlDoc string) (*zip.File, error) {
	c := 0
	f := new(zip.File)
	for _, file := range files.File {
		if file.Name == xmlDoc {
			f = file
			c++
		}
	}

	if c == 0 {
		return nil, fmt.Errorf("no %s found in docx file", xmlDoc)
	} else if c > 1 {
		return nil, fmt.Errorf("more than one %s found in docx file", xmlDoc)
	}

	return f, nil
}

//openFile : Opens the zipfile into a reader
func openFile(file zip.File) (io.ReadCloser, error) {
	rc, err := file.Open()
	if err != nil {
		return nil, err
	}
	return rc, nil
}

//fileToString : This converts the file into a string
func fileToString(reader io.Reader) (string, error) {
	content, err := io.ReadAll(reader)
	if err != nil {
		return "", err
	}
	return string(content), nil
}

//extractDocument : parse the string of xml into the structure defined as "Document"
func extractDocument(xmlContent string) (Document, error) {
	s := Document{}
	err := xml.Unmarshal([]byte(xmlContent), &s)
	if err != nil {
		return Document{}, err
	}
	return s, nil
}

//extractHeader : parse the string of xml into the structure defined as "Header"
func extractHeader(xmlContent string) (Header, error) {
	s := Header{}
	err := xml.Unmarshal([]byte(xmlContent), &s)
	if err != nil {
		return Header{}, err
	}
	return s, nil
}

//extractFooter : parse the string of xml into the structure defined as "Footer"
func extractFooter(xmlContent string) (Footer, error) {
	s := Footer{}
	err := xml.Unmarshal([]byte(xmlContent), &s)
	if err != nil {
		return Footer{}, err
	}
	return s, nil
}

//GetDocument : takes a file, and retuns the Document{} struct
func GetDocument(f string) (Document, error) {
	reader, err := unzipDocx(f)
	if err != nil {
		return Document{}, err
	}
	defer reader.Close()

	file, err := getXMLFile(reader, docxPathOfDocumentXML)
	if err != nil {
		return Document{}, err
	}

	fileContent, err := openFile(*file)
	if err != nil {
		return Document{}, err
	}
	defer fileContent.Close()

	content, err := fileToString(fileContent)
	if err != nil {
		return Document{}, err
	}

	d, err := extractDocument(content)
	if err != nil {
		return Document{}, err
	}
	return d, nil
}

//GetHeader : takes a file, and retuns the Header{} struct
func GetHeader(f string) (Header, error) {
	reader, err := unzipDocx(f)
	if err != nil {
		return Header{}, err
	}
	defer reader.Close()

	file, err := getXMLFile(reader, docxPathOfHeaderXML)
	if err != nil {
		return Header{}, err
	}

	fileContent, err := openFile(*file)
	if err != nil {
		return Header{}, err
	}
	defer fileContent.Close()

	content, err := fileToString(fileContent)
	if err != nil {
		return Header{}, err
	}

	hdr, err := extractHeader(content)
	if err != nil {
		return Header{}, err
	}
	return hdr, nil
}

//GetFooter : takes a file, and retuns the Footer{} struct
func GetFooter(f string) (Footer, error) {
	reader, err := unzipDocx(f)
	if err != nil {
		return Footer{}, err
	}
	defer reader.Close()

	file, err := getXMLFile(reader, docxPathOfFooterXML)
	if err != nil {
		return Footer{}, err
	}

	fileContent, err := openFile(*file)
	if err != nil {
		return Footer{}, err
	}
	defer fileContent.Close()

	content, err := fileToString(fileContent)
	if err != nil {
		return Footer{}, err
	}

	ftr, err := extractFooter(content)
	if err != nil {
		return Footer{}, err
	}
	return ftr, nil
}
