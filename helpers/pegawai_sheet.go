package helpers

import (
	"github.com/xuri/excelize/v2"
	"log"
)

const (
	PegawaiSheetName        = "Pegawai"
	PegawaiDataStartRowFrom = 5
	PegawaiColumnNo         = 1
	PegawaiColumnName       = 2
	PegawaiColumnCode       = 3
	PegawaiColumnGol        = 4
	PegawaiColumnNIP        = 5
	PegawaiColumnNJabatan   = 7
	PegawaiColumnSeksiUtama = 8
	PegawaiEndColumn        = PegawaiColumnSeksiUtama

	KesepakatanTargetCell = "K5"
)

func GetDetailPegawai(file *excelize.File) (details []map[string]string) {

	// get all rows
	rows, err := file.GetRows(PegawaiSheetName)
	if err != nil {
		log.Fatal(err.Error())
	}

	// get details
	for i, val := range rows {
		if i >= PegawaiDataStartRowFrom && (len(val) >= PegawaiEndColumn && val[PegawaiColumnNo] != "") {
			detailTemp := map[string]string{}
			detailTemp["nama"] = val[PegawaiColumnName]
			detailTemp["code"] = val[PegawaiColumnCode]
			detailTemp["gol"] = val[PegawaiColumnGol]
			detailTemp["nip"] = val[PegawaiColumnNIP]
			detailTemp["jabatan"] = val[PegawaiColumnNJabatan]
			detailTemp["seksi_utama"] = val[PegawaiColumnSeksiUtama]
			details = append(details, detailTemp)
		}
	}

	return details
}

func GetKesepakatanTarget(file *excelize.File) string {

	// get all rows
	value, _ := file.GetCellValue(PegawaiSheetName, KesepakatanTargetCell)

	return value
}
