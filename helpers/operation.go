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
	CKTPKSKException            = "KSK"

	StyleArial            = `{"font":{"family": "Arial"}}`
	StyleBorder           = `{"border": [{"type": "left","color": "000000","style": 1},{"type": "top","color": "000000","style": 1},{"type": "bottom","color": "000000","style": 1},{"type": "right","color": "000000","style": 1}]}`
	StyleBoldAndUnderline = `{"font":{"bold": true, "underline":"single", "size": 11, "family": "Arial"}}`
	StyleBorderCenter     = `{"alignment":{"horizontal": "center"}, "border": [{"type": "left","color": "000000","style": 1},{"type": "top","color": "000000","style": 1},{"type": "bottom","color": "000000","style": 1},{"type": "right","color": "000000","style": 1}]}`

	StyleJumlah  = `{"font":{"family": "Arial"}, "font":{"bold": true}, "alignment":{"horizontal": "center", "vertical": "center"}, "border": [{"type": "left","color": "000000","style": 1},{"type": "top","color": "000000","style": 1},{"type": "bottom","color": "000000","style": 6},{"type": "right","color": "000000","style": 1}], "fill":{"type":"pattern","color":["#BFBFBF"],"pattern":1}}`
	StylePenilai = `{"font":{"family": "Segoe UI", "size": 10}, "alignment":{"horizontal": "center", "vertical": "center"}}`
	StyleNamaPenilai = `{"font":{"family": "Segoe UI", "size": 10, "bold": true, "underline":"single"}, "alignment":{"horizontal": "center", "vertical": "center"}}`
	StyleKesepakatan = `{"font":{"family": "Arial", "bold": true}}`
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

func WriteSpecificCKPT(file *excelize.File, mergeData []map[string]interface{}, kesepakatanTarget string) {

	// write every sheet
	for _, data := range mergeData {

		// init
		var columnNumber int
		var counter int
		sheetName := fmt.Sprintf("%s %s", CKPTTemplatePrefixSheetName, data["code"])

		// style
		arial, _ := file.NewStyle(StyleArial)
		border, _ := file.NewStyle(StyleBorder)
		boldAndUnderline, _ := file.NewStyle(StyleBoldAndUnderline)
		borderCenter, _ := file.NewStyle(StyleBorderCenter)
		jumlah, _ := file.NewStyle(StyleJumlah)
		penilai, _ := file.NewStyle(StylePenilai)
		namaPenilai, _ := file.NewStyle(StyleNamaPenilai)
		kesepakatan, _ := file.NewStyle(StyleKesepakatan)

		// default font family
		file.SetCellStyle(sheetName, "A12", "I100", arial)

		// write variable
		namaLengkapPegawai := fmt.Sprintf("Nama%23s: %s", " ", data["nama"])
		jabatanPegawai := fmt.Sprintf("Jabatan%20s: %s", " ", data["jabatan"])

		// write
		file.SetCellValue(sheetName, "B4", CKPTSatuanOrganisasi)
		file.SetCellValue(sheetName, "B5", namaLengkapPegawai)
		file.SetCellValue(sheetName, "B6", jabatanPegawai)
		file.SetCellValue(sheetName, "B7", CKPTPeriode)

		// Utama
		counter = 0
		columnNumber = 0
		file.SetCellValue(sheetName, "A12", "UTAMA")
		file.SetCellStyle(sheetName, "A12", "A12", boldAndUnderline)
		file.SetCellStyle(sheetName, "B12", "I12", border)
		for _, kegiatan := range data["kegiatan"].([]map[string]string) {
			if kegiatan["fungsi"] == data["seksi_utama"] || data["seksi_utama"] == CKTPKSKException {
				columnNumber += 1

				// write
				file.SetCellValue(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), columnNumber)
				file.SetCellValue(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter), kegiatan["kegiatan"])
				file.SetCellValue(sheetName, fmt.Sprintf("C%d", CKPTStartWriteKegiatan+counter), kegiatan["satuan"])
				file.SetCellValue(sheetName, fmt.Sprintf("D%d", CKPTStartWriteKegiatan+counter), kegiatan["target"])

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
		columnNumber = 0
		file.SetCellValue(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), "TAMBAHAN")
		file.SetCellStyle(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), boldAndUnderline)
		file.SetCellStyle(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter), fmt.Sprintf("I%d", CKPTStartWriteKegiatan+counter), border)
		counter += 1
		for _, kegiatan := range data["kegiatan"].([]map[string]string) {
			if kegiatan["fungsi"] != data["seksi_utama"] && data["seksi_utama"] != CKTPKSKException {
				columnNumber += 1

				// write
				file.SetCellValue(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), columnNumber)
				file.SetCellValue(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter), kegiatan["kegiatan"])
				file.SetCellValue(sheetName, fmt.Sprintf("C%d", CKPTStartWriteKegiatan+counter), kegiatan["satuan"])
				file.SetCellValue(sheetName, fmt.Sprintf("D%d", CKPTStartWriteKegiatan+counter), kegiatan["target"])

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

		// ------------------------------------------------- Summary ------------------------------------------------------- //

		// jumlah
		file.SetCellValue(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), "JUMLAH")
		file.MergeCell(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter))
		file.SetCellStyle(sheetName, fmt.Sprintf("A%d", CKPTStartWriteKegiatan+counter), fmt.Sprintf("G%d", CKPTStartWriteKegiatan+counter), jumlah)

		// kesepakatan
		file.SetCellValue(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter+2), "Kesepakatan Target")
		file.SetCellStyle(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter+2), fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter+2), kesepakatan)
		file.SetCellValue(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter+3), kesepakatanTarget)

		// pejabat dinilai
		file.SetCellValue(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter+4), "Pegawai Yang Dinilai")
		file.SetCellStyle(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter+4), fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter+4), penilai)

		file.SetCellValue(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter+7), data["nama"])
		file.SetCellStyle(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter+7), fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter+7), namaPenilai)

		file.SetCellValue(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter+8), fmt.Sprintf("NIP. %s", data["nip"]))
		file.SetCellStyle(sheetName, fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter+8), fmt.Sprintf("B%d", CKPTStartWriteKegiatan+counter+8), penilai)


		// pejabat penilai
		file.MergeCell(sheetName, fmt.Sprintf("E%d", CKPTStartWriteKegiatan+counter+4), fmt.Sprintf("H%d", CKPTStartWriteKegiatan+counter+4))
		file.SetCellValue(sheetName, fmt.Sprintf("E%d", CKPTStartWriteKegiatan+counter+4), "Pejabat Penilai")
		file.SetCellStyle(sheetName, fmt.Sprintf("E%d", CKPTStartWriteKegiatan+counter+4), fmt.Sprintf("E%d", CKPTStartWriteKegiatan+counter+4), penilai)

		file.MergeCell(sheetName, fmt.Sprintf("E%d", CKPTStartWriteKegiatan+counter+7), fmt.Sprintf("H%d", CKPTStartWriteKegiatan+counter+7))
		file.SetCellValue(sheetName, fmt.Sprintf("E%d", CKPTStartWriteKegiatan+counter+7), CKPTPejabatPenilai)
		file.SetCellStyle(sheetName, fmt.Sprintf("E%d", CKPTStartWriteKegiatan+counter+7), fmt.Sprintf("E%d", CKPTStartWriteKegiatan+counter+7), namaPenilai)

		file.MergeCell(sheetName, fmt.Sprintf("E%d", CKPTStartWriteKegiatan+counter+8), fmt.Sprintf("H%d", CKPTStartWriteKegiatan+counter+8))
		file.SetCellValue(sheetName, fmt.Sprintf("E%d", CKPTStartWriteKegiatan+counter+8), CKTPPejabatPenilaiNIP)
		file.SetCellStyle(sheetName, fmt.Sprintf("E%d", CKPTStartWriteKegiatan+counter+8), fmt.Sprintf("E%d", CKPTStartWriteKegiatan+counter+8), penilai)
		// ------------------------------------------------------------------------------------------------------------------ //


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
