package main

import (
	"oqxa/helpers"
)

func main() {

	// init variables
	file := helpers.OpenExcel()
	names := helpers.GetNames(file)
	detailPegawai := helpers.GetDetailPegawai(file)
	detailKegiatan := helpers.GetDetailKegiatan(file, names)
	mergeData := mergeKegiatanDanPegawai(detailPegawai, detailKegiatan)

	// create sheet and copy template
	indexes := helpers.CreateSheets(file, names)
	helpers.CopyTemplate(file, indexes)

	// write all sheet
	helpers.WriteSpecificCKPT(file, mergeData)

	// save template
	helpers.SaveAsFile(file)
}

func mergeKegiatanDanPegawai(pegawai []map[string]string, kegiatan []map[string]interface{}) (merged []map[string]interface{}) {
	for _, peg := range pegawai {
		for _, keg := range kegiatan {
			if peg["code"] == keg["code"] {
				mergedTemp := map[string]interface{}{
					"code": peg["code"],
					"nama": peg["nama"],
					"nip": peg["nip"],
					"jabatan": peg["jabatan"],
					"seksi_utama": peg["seksi_utama"],
					"kegiatan": keg["kegiatan"],
				}
				merged = append(merged, mergedTemp)
			}
		}
	}

	return merged
}
