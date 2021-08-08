package helpers

import (
	"github.com/xuri/excelize/v2"
	"log"
)

const(
	CKPTAllSourceFilePath   = "Template.xlsx"
)

func OpenExcel() (file *excelize.File) {
	file, err := excelize.OpenFile(CKPTAllSourceFilePath)
	if err != nil {
		log.Fatal(err.Error())
	}

	return file
}
