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

var (
	TEMPLATE_HEADER = "Code4Func"
)

func init() {

}
func main() {
	initReport()
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
	orderReportSheetName := "report"
	index := f.NewSheet(orderReportSheetName)

	f.MergeCell(orderReportSheetName, "D2", "G2")
	f.SetCellValue(orderReportSheetName, "D2", header.CompanyName)

	f.MergeCell(orderReportSheetName, "D3", "G3")
	f.SetCellValue(orderReportSheetName, "D3", header.CompanyAddress)

	f.MergeCell(orderReportSheetName, "D4", "G4")
	f.SetCellValue(orderReportSheetName, "D4", header.ContactInfo)

	f.MergeCell(orderReportSheetName, "L3", "N3")
	f.SetCellValue(orderReportSheetName, "L3", header.ReportDate)

	f.MergeCell(orderReportSheetName, "L4", "N4")
	f.SetCellValue(orderReportSheetName, "L4", header.InvoiceId)

	f.SetActiveSheet(index)
	if err := f.SaveAs("report.xlsx"); err != nil {
		fmt.Println(err)
	}
}

func mergeAndSaveValueCell(file *excelize.File, nameOfSheet, beginCell, end string,
	value interface{}) {
	f.MergeCell(nameOfSheet, beginRow, endRow)
	f.SetCellValue(nameOfSheet, D2, D2)
}
