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
	CKPTPejabatPenilai          = "Tony Suprianto, S.Si., M.A.P."
	CKTPPejabatPenilaiNIP       = "NIP. 197605212002121002"

	StyleArial            = `{"font":{"family": "Arial"}}`
	StyleBorder           = `{"border": [{"type": "left","color": "000000","style": 1},{"type": "top","color": "000000","style": 1},{"type": "bottom","color": "000000","style": 1},{"type": "right","color": "000000","style": 1}]}`
	StyleBoldAndUnderline = `{"font":{"bold": true, "underline":"single", "size": 11, "family": "Arial"}}`
	StyleCenter           = `{"alignment":{"horizontal": "center"}}`
	StyleBorderCenter     = `{"alignment":{"horizontal": "center"}, "border": [{"type": "left","color": "000000","style": 1},{"type": "top","color": "000000","style": 1},{"type": "bottom","color": "000000","style": 1},{"type": "right","color": "000000","style": 1}]}`

	StyleJumlah = `{"font":{"family": "Arial"}, "font":{"bold": true}, "alignment":{"horizontal": "center", "vertical": "center"}, "border": [{"type": "left","color": "000000","style": 1},{"type": "top","color": "000000","style": 1},{"type": "bottom","color": "000000","style": 6},{"type": "right","color": "000000","style": 1}], "fill":{"type":"pattern","color":["#BFBFBF"],"pattern":1}}`
)

var (
	CKPTSatuanOrganisasi = fmt.Sprintf("Satuan Organisasi%3s: %s", " ", "BPS Kabupaten Lamandau")
	CKPTPeriode          = fmt.Sprintf("Periode%20s: %s", " ", "")
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
		arial, _ := file.NewStyle(StyleArial)
		border, _ := file.NewStyle(StyleBorder)
		boldAndUnderline, _ := file.NewStyle(StyleBoldAndUnderline)
		//center, _ := file.NewStyle(StyleCenter)
		borderCenter, _ := file.NewStyle(StyleBorderCenter)
		jumlah, _ := file.NewStyle(StyleJumlah)

		// sheet name
		sheetName := fmt.Sprintf("%s %s", CKPTTemplatePrefixSheetName, val["nama"])

		// sheet font family
		file.SetCellStyle(sheetName, "A12", "I100", arial)

		for _, detailPegawai := range pegawai {
			if detailPegawai["code"] == val["nama"] {

				// write variable
				namaLengkapPegawai := fmt.Sprintf("Nama%23s: %s", " ", detailPegawai["nama"])
				jabatanPegawai := fmt.Sprintf("Jabatan%20s: %s", " ", detailPegawai["jabatan"])

				// write
				file.SetCellValue(sheetName, "B4", CKPTSatuanOrganisasi)
				file.SetCellValue(sheetName, "B5", namaLengkapPegawai)
				file.SetCellValue(sheetName, "B6", jabatanPegawai)
				file.SetCellValue(sheetName, "B7", CKPTPeriode)

				// Utama
				file.SetCellValue(sheetName, "A12", "UTAMA")
				file.SetCellStyle(sheetName, "A12", "A12", boldAndUnderline)
				file.SetCellStyle(sheetName, "B12", "I12", border)

				// write kegiatan
				counter = 0
				columnNumber = 0
				for _, detailKegiatan := range val["kegiatan"].([]map[string]string) {
					if detailKegiatan["fungsi"] == detailPegawai["seksi_utama"] {

						columnNumber += 1

						// write
						file.SetCellValue(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), columnNumber)
						file.SetCellValue(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter), detailKegiatan["kegiatan"])
						file.SetCellValue(sheetName, fmt.Sprintf("C%d", CKPTStartWriteKegiatan+counter), detailKegiatan["satuan"])
						file.SetCellValue(sheetName, fmt.Sprintf("D%d", CKPTStartWriteKegiatan+counter), detailKegiatan["target"])

						// add border
						file.SetCellStyle(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), fmt.Sprintf("I%d", CKPTStartWriteKegiatan+counter), border)
						file.SetCellStyle(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), borderCenter)
						file.SetCellStyle(sheetName, fmt.Sprintf("C%d", CKPTStartWriteKegiatan+counter), fmt.Sprintf("I%d", CKPTStartWriteKegiatan+counter), borderCenter)

						counter += 1

					}
				}

				// add cell - if 0
				if columnNumber == 0 {
					file.SetCellValue(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), columnNumber+1)
					file.SetCellStyle(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), fmt.Sprintf("I%d", CKPTStartWriteKegiatan+counter), border)
					file.SetCellStyle(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), borderCenter)
					file.SetCellStyle(sheetName, fmt.Sprintf("C%d", CKPTStartWriteKegiatan+counter), fmt.Sprintf("I%d", CKPTStartWriteKegiatan+counter), borderCenter)
					counter++
				}

				// Tambahan
				file.SetCellValue(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), "TAMBAHAN")
				file.SetCellStyle(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), boldAndUnderline)
				file.SetCellStyle(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter), fmt.Sprintf("I%d", CKPTStartWriteKegiatan+counter), border)
				counter += 1

				// seksi tambahan
				columnNumber = 0
				for _, detailKegiatan := range val["kegiatan"].([]map[string]string) {
					if detailKegiatan["fungsi"] != detailPegawai["seksi_utama"] {

						columnNumber += 1

						// write
						file.SetCellValue(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), columnNumber)
						file.SetCellValue(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter), detailKegiatan["kegiatan"])
						file.SetCellValue(sheetName, fmt.Sprintf("C%d", CKPTStartWriteKegiatan+counter), detailKegiatan["satuan"])
						file.SetCellValue(sheetName, fmt.Sprintf("D%d", CKPTStartWriteKegiatan+counter), detailKegiatan["target"])

						// add border
						file.SetCellStyle(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), fmt.Sprintf("I%d", CKPTStartWriteKegiatan+counter), border)
						file.SetCellStyle(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), borderCenter)
						file.SetCellStyle(sheetName, fmt.Sprintf("C%d", CKPTStartWriteKegiatan+counter), fmt.Sprintf("I%d", CKPTStartWriteKegiatan+counter), borderCenter)

						counter += 1

					}
				}

				// add cell - if 0
				if columnNumber == 0 {
					file.SetCellValue(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), columnNumber+1)
					file.SetCellStyle(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), fmt.Sprintf("I%d", CKPTStartWriteKegiatan+counter), border)
					file.SetCellStyle(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), borderCenter)
					file.SetCellStyle(sheetName, fmt.Sprintf("C%d", CKPTStartWriteKegiatan+counter), fmt.Sprintf("I%d", CKPTStartWriteKegiatan+counter), borderCenter)
					counter++
				}

				// summary
				file.SetCellValue(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), "JUMLAH")
				file.MergeCell(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter))
				file.SetCellStyle(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), fmt.Sprintf("G%d", CKPTStartWriteKegiatan+counter), jumlah)

				file.SetCellValue(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter+2), "Kesepakatan Target")
				file.SetCellValue(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter+3), "Tanggal:")
				file.SetCellValue(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter+3), "Pegawai Yang Dinilai")
				file.SetCellValue(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter+6), detailPegawai["nama"])
				file.SetCellValue(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter+7), fmt.Sprintf("NIP. %s", detailPegawai["nip"]))
				file.SetCellValue(sheetName, fmt.Sprintf("E%d", CKPTStartWriteKegiatan+counter+3), "Pejabat Penilai")
				file.SetCellValue(sheetName, fmt.Sprintf("E%d", CKPTStartWriteKegiatan+counter+6), CKPTPejabatPenilai)
				file.SetCellValue(sheetName, fmt.Sprintf("E%d", CKPTStartWriteKegiatan+counter+7), CKTPPejabatPenilaiNIP)

			}
		}

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
