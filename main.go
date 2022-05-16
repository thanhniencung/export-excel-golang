package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

type HeaderReport struct {
	CompanyName    string
	CompanyAddress string
	ContactInfo    string
	ReportDate     string
	InvoiceId      string
}

func main() {
	initReport()
}

var orderReportSheetName = "report"

func setCellValue(f *excelize.File, axis string, value interface{}) {
	f.SetCellValue(orderReportSheetName, axis, value)
}

func initReport() {
	f := excelize.NewFile()

	header := HeaderReport{
		CompanyName:    "Code4Func",
		CompanyAddress: "Quang Trung SoftWare Park",
		ContactInfo:    "0123456789 | www.code4func.com",
		ReportDate:     "16/05/2022",
		InvoiceId:      "#001",
	}

	index := f.NewSheet(orderReportSheetName)

	for i := 2; i<5; i++ {
		f.MergeCell(orderReportSheetName, fmt.Sprintf("D%d", i), fmt.Sprintf("G%d", i))
	}

	for i := 3; i<5; i++ {
		f.MergeCell(orderReportSheetName, fmt.Sprintf("L%d", i), fmt.Sprintf("N%d", i))
	}

	mapHeaderData := map[string] string {
		"D2": header.CompanyName,
		"D3": header.CompanyAddress,
 		"D4": header.ContactInfo,
		"L3": header.ReportDate,
		"L4": header.InvoiceId,
	}

	for key, value := range mapHeaderData {
		setCellValue(f, key, value)
	}

	f.SetActiveSheet(index)
	if err := f.SaveAs("report.xlsx"); err != nil {
		fmt.Println(err)
	}
}


