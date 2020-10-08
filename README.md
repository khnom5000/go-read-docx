# go-read-docx
simple way to read a docx in go


	reader, readerFiles := docx.UnpackDocx("./TestDocument.docx")
	defer reader.Close()
	docFile := docx.RetrieveWordDoc(readerFiles)
	doc := docx.OpenWordDoc(*docFile)
	content := docx.WordDocToString(doc)
	fmt.Print(docx.Extract(content))

As a function

	func readDocxTable(filename string) (table [][]string) {
		reader, readerFiles := docx.UnpackDocx(filename)
		defer reader.Close()
		docFile := docx.RetrieveWordDoc(readerFiles)
		doc := docx.OpenWordDoc(*docFile)
		content := docx.WordDocToString(doc)
		t := docx.Extract(content).Body.Table.TableRow
	
		for _, r := range t {
			var row []string
			for _, c := range r.TableColumn {
				row = append(row, strings.Join(c.Cell, ""))
			}
			table = append(table, row)
		}
		return
	}
