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

var f *excelize.File
func init() {
	f = excelize.NewFile()
}

func main() {
	initReport()
	export(0)
}

var orderReportSheetName = "report"

func setCellValue(axis string, value interface{}) {
	f.SetCellValue(orderReportSheetName, axis, value)
}

func buildHeaderReport() {
	header := HeaderReport{
		CompanyName:    "Code4Func",
		CompanyAddress: "Quang Trung SoftWare Park",
		ContactInfo:    "0123456789 | www.code4func.com",
		ReportDate:     "16/05/2022",
		InvoiceId:      "#001",
	}

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
		setCellValue(key, value)
	}
}

func buildTableHeader() {
	cols := []string{ "D", "E", "F", "G", "H"}
	colLabel := []string{ "Production ID", "Description", "Quantity", "Unit price", "Total"}

	for index, col := range cols {
		setCellValue(fmt.Sprintf("%s%d", col, index + 7), colLabel[index])
	}
}

type Order struct {
	ProductID string
	Description string
	Quantity int
	UnitPrice float64
	Total int
}

func buildFakeOrder() []Order {
	var orders []Order
	for i:=0; i< 20; i++ {
		orders = append(orders, Order{
			ProductID:   fmt.Sprintf("#p%d", i),
			Description: fmt.Sprintf("#Item %d", i),
			Quantity:    i,
			UnitPrice:   200,
			Total:       200 * i,
		})
	}

	return orders
}

func buildTableContent() {
	orders := buildFakeOrder()
	for index, order := range orders {
		setCellValue(fmt.Sprintf("D%d", index + 7 ), order.ProductID)
		setCellValue(fmt.Sprintf("E%d", index + 7 ), order.Description)
		setCellValue(fmt.Sprintf("F%d", index + 7 ), order.Quantity)
		setCellValue(fmt.Sprintf("G%d", index + 7 ), order.UnitPrice)
		setCellValue(fmt.Sprintf("H%d", index + 7 ), order.Total)
	}
}


func initReport() {
	buildHeaderReport()
	buildTableHeader()
	buildTableContent()
}

func export(index int) {
	f.SetActiveSheet(index)
	if err := f.SaveAs("report.xlsx"); err != nil {
		fmt.Println(err)
	}
}


