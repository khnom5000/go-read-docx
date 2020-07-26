# go-read-docx
simple way to read a docx in go


	reader, readerFiles := docx.UnpackDocx("./TestDocument.docx")
	defer reader.Close()
	docFile := docx.RetrieveWordDoc(readerFiles)
	doc := docx.OpenWordDoc(*docFile)
	content := docx.WordDocToString(doc)
	fmt.Print(docx.Extract(content))
