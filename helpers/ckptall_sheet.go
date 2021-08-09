package helpers

import (
	"github.com/xuri/excelize/v2"
	"log"
)

const (
	CKPTAllSheetName        = "CKP-T All"
	CKPTAllDataStartRowFrom = 5
	CKPTAllColumnName       = 0
	CKPTAllColumnFungsi     = 1
	CKPTAllColumnKegiatan   = 2
	CKPTAllColumnSatuan     = 3
	CKPTAllColumnTarget     = 4
)

func GetNames(file *excelize.File) (names []string) {

	// get all rows
	rows, err := file.GetRows(CKPTAllSheetName)
	if err != nil {
		log.Fatal(err.Error())
	}

	// get names
	for i, val := range rows {
		if i >= CKPTAllDataStartRowFrom {
			isExist := false
			for _, name := range names {
				if name == val[CKPTAllColumnName] {
					isExist = true
					break
				}
			}
			if !isExist {
				names = append(names, val[CKPTAllColumnName])
			}
		}
	}

	return names
}

func GetDetailKegiatan(file *excelize.File, names []string) (details []map[string]interface{}) {

	// get all rows
	rows, err := file.GetRows(CKPTAllSheetName)
	if err != nil {
		log.Fatal(err.Error())
	}

	// get all detail
	for _, name := range names {
		detail := map[string]interface{}{} // {name: adam, details: [{}]}
		detail["code"] = name

		var detailKegiatan []map[string]string
		for i, val := range rows {
			if i >= CKPTAllDataStartRowFrom {
				if name == val[CKPTAllColumnName] {
					kegiatanTemp := map[string]string{
						"fungsi":   val[CKPTAllColumnFungsi],
						"kegiatan": val[CKPTAllColumnKegiatan],
						"satuan":   val[CKPTAllColumnSatuan],
						"target":   val[CKPTAllColumnTarget],
					}
					detailKegiatan = append(detailKegiatan, kegiatanTemp)
				}
			}
		}
		detail["kegiatan"] = detailKegiatan
		details = append(details, detail)
	}

	return details
}
