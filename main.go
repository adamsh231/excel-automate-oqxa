package main

import (
	"oqxa/helpers"
)

func main() {

	// init variables
	file := helpers.OpenExcel()
	names := helpers.GetNames(file)
	detailKegiatan := helpers.GetDetailKegiatan(file, names)
	detailPegawai := helpers.GetDetailPegawai(file)

	// create sheet and copy template
	indexes := helpers.CreateSheets(file, names)
	helpers.CopyTemplate(file, indexes)

	// write all sheet
	helpers.WriteSpecificCKPT(file, detailPegawai, detailKegiatan)

	// save template
	helpers.SaveAsFile(file)

	//fmt.Println(file.GetSheetList())
}
