package main

import (
	"fmt"
	"github.com/unidoc/unioffice/document"
)

func main() {
	doc, err := document.Open("/Users/kadyrbeknarmamatov/go/src/github.com/SEFI2/unioffice-run/watermark.docx")
	if err != nil {
		fmt.Println(err)
		return
	}

	para := doc.AddParagraph()
	run := para.AddRun()
	run.AddText("Whatever, this is just test text")

	if err := doc.SaveToFile("/Users/kadyrbeknarmamatov/go/src/github.com/SEFI2/unioffice-run/result.docx"); err != nil {
		fmt.Println(err)
		return
	}
}
