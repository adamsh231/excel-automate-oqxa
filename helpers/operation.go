package helpers

import (
	"fmt"
	"github.com/google/uuid"
	"github.com/xuri/excelize/v2"
	"log"
)

const (
	CKPTTemplateSheetName       = "CKP-T Template"
	CKPTTemplatePrefixSheetName = "CKP-T"
	CKPTStartWriteKegiatan      = 13
)

func CreateSheets(file *excelize.File, names []string) (indexes []int) {

	// create sheet based on names
	for _, name := range names {
		sheetName := fmt.Sprintf("%s %s", CKPTTemplatePrefixSheetName, name)
		file.NewSheet(sheetName)
		indexes = append(indexes, file.GetSheetIndex(sheetName))
	}

	return indexes
}

func CopyTemplate(file *excelize.File, indexes []int) {

	// copying template based on indexes
	for _, index := range indexes {
		if err := file.CopySheet(file.GetSheetIndex(CKPTTemplateSheetName), index); err != nil {
			log.Fatal(err.Error())
		}
	}

}

func WriteSpecificCKPT(file *excelize.File, pegawai []map[string]string, kegiatan []map[string]interface{}) {

	// write every sheet
	for _, val := range kegiatan {

		// init column number
		var columnNumber int
		var counter int

		// sheet name
		sheetName := fmt.Sprintf("%s %s", CKPTTemplatePrefixSheetName, val["nama"])

		// sheet font family
		arial, _ := file.NewStyle(`{"font":{"family": "Arial"}}`)
		file.SetCellStyle(sheetName, "A12", "I100", arial)

		for _, detailPegawai := range pegawai {
			if detailPegawai["code"] == val["nama"] {

				// write variable
				namaLengkapPegawai := fmt.Sprintf("Nama%23s: %s", " ", detailPegawai["nama"])
				jabatanPegawai := fmt.Sprintf("Jabatan%20s: %s", " ", detailPegawai["jabatan"])

				// write
				file.SetCellValue(sheetName, "B5", namaLengkapPegawai)
				file.SetCellValue(sheetName, "B6", jabatanPegawai)

				// Utama
				file.SetCellValue(sheetName, "A12", "UTAMA")
				boldAndUnderline, _ := file.NewStyle(`{"font":{"bold": true, "size": 11, "family": "Arial"}}`)
				file.SetCellStyle(sheetName, "A12", "A12", boldAndUnderline)

				// write kegiatan
				counter = 0
				columnNumber = 0
				for _, detailKegiatan := range val["kegiatan"].([]map[string]string) {
					if detailKegiatan["fungsi"] == detailPegawai["seksi_utama"] {

						// write
						columnNumber += 1
						file.SetCellValue(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), columnNumber)
						file.SetCellValue(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter), detailKegiatan["kegiatan"])
						file.SetCellValue(sheetName, fmt.Sprintf("C%d", CKPTStartWriteKegiatan+counter), detailKegiatan["satuan"])
						file.SetCellValue(sheetName, fmt.Sprintf("D%d", CKPTStartWriteKegiatan+counter), detailKegiatan["target"])
						counter += 1

					}
				}

				// add cell - if 0
				if columnNumber == 0 {
					file.SetCellValue(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), "1")
					file.SetCellValue(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter), "-")
					counter++
				}

				// Tambahan
				file.SetCellValue(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), "TAMBAHAN")
				file.SetCellStyle(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), boldAndUnderline)
				counter += 1

				// seksi tambahan
				columnNumber = 0
				for _, detailKegiatan := range val["kegiatan"].([]map[string]string) {
					if detailKegiatan["fungsi"] != detailPegawai["seksi_utama"] {

						// write
						columnNumber += 1
						file.SetCellValue(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), columnNumber)
						file.SetCellValue(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter), detailKegiatan["kegiatan"])
						file.SetCellValue(sheetName, fmt.Sprintf("C%d", CKPTStartWriteKegiatan+counter), detailKegiatan["satuan"])
						file.SetCellValue(sheetName, fmt.Sprintf("D%d", CKPTStartWriteKegiatan+counter), detailKegiatan["target"])
						counter += 1

					}
				}

				// add cell - if 0
				if columnNumber == 0 {
					file.SetCellValue(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), "1")
					file.SetCellValue(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter), "-")
					counter++
				}

			}
		}

		border, _ := file.NewStyle(`{"border": [{"type": "left","color": "000000","style": 1},{"type": "top","color": "000000","style": 1},{"type": "bottom","color": "000000","style": 1},{"type": "right","color": "000000","style": 1}]}`)
		file.SetCellStyle(sheetName, "A12", fmt.Sprintf("I%d", CKPTStartWriteKegiatan+counter), border)
	}

}

func SaveFile(file *excelize.File) {

	// save file
	if err := file.Save(); err != nil {
		log.Fatal(err.Error())
	}

}

func SaveAsFile(file *excelize.File) {

	// save file
	if err := file.SaveAs("result_" + uuid.New().String() + ".xlsx"); err != nil {
		log.Fatal(err.Error())
	}

}
