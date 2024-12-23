package main

import (
	"flag"
	"fmt"
	"github.com/xuri/excelize/v2"
	"time"
)

const outputSheetName = "Sheet1"

var (
	input          = flag.String("input", "input.xlsx", "input file")
	inputSheetName = flag.String("inputSheet", "Sheet1", "input sheet name")
	output         = flag.String("output", "output.xlsx", "output file")
)

type Row struct {
	Date                    string
	Category                string
	User                    string
	Task                    string
	State                   string
	Time                    string
	Region                  string
	Country                 string
	Department              string
	OrganizationalHierarchy string
	FunctionalArea          string
	ServiceCatalog          string
	ServiceOffering         string
}

func createHeader(outputFile *excelize.File) {
	outputFile.SetCellStr(outputSheetName, "A1", "Date")
	outputFile.SetCellStr(outputSheetName, "B1", "Category")
	outputFile.SetCellStr(outputSheetName, "C1", "User")
	outputFile.SetCellStr(outputSheetName, "D1", "Task")
	outputFile.SetCellStr(outputSheetName, "E1", "Time")
	outputFile.SetCellStr(outputSheetName, "F1", "Region")
	outputFile.SetCellStr(outputSheetName, "G1", "Country")
	outputFile.SetCellStr(outputSheetName, "H1", "Department")
	outputFile.SetCellStr(outputSheetName, "I1", "Organizational Hierarchy")
	outputFile.SetCellStr(outputSheetName, "J1", "Functional Area")
	outputFile.SetCellStr(outputSheetName, "K1", "Service catalog")
	outputFile.SetCellStr(outputSheetName, "L1", "Service offering")
}

func writeRows(outputFile *excelize.File, rows []Row, rowIndex *int) {
	*rowIndex = *rowIndex + 1
	for _, row := range rows {
		outputFile.SetCellStr(outputSheetName, fmt.Sprintf("A%d", *rowIndex), row.Date)
		outputFile.SetCellStr(outputSheetName, fmt.Sprintf("B%d", *rowIndex), row.Category)
		outputFile.SetCellStr(outputSheetName, fmt.Sprintf("C%d", *rowIndex), row.User)
		outputFile.SetCellStr(outputSheetName, fmt.Sprintf("D%d", *rowIndex), row.Task)
		outputFile.SetCellStr(outputSheetName, fmt.Sprintf("E%d", *rowIndex), row.Time)
		outputFile.SetCellStr(outputSheetName, fmt.Sprintf("F%d", *rowIndex), row.Region)
		outputFile.SetCellStr(outputSheetName, fmt.Sprintf("G%d", *rowIndex), row.Country)
		outputFile.SetCellStr(outputSheetName, fmt.Sprintf("H%d", *rowIndex), row.Department)
		outputFile.SetCellStr(outputSheetName, fmt.Sprintf("I%d", *rowIndex), row.OrganizationalHierarchy)
		outputFile.SetCellStr(outputSheetName, fmt.Sprintf("J%d", *rowIndex), row.FunctionalArea)
		outputFile.SetCellStr(outputSheetName, fmt.Sprintf("K%d", *rowIndex), row.ServiceCatalog)
		outputFile.SetCellStr(outputSheetName, fmt.Sprintf("L%d", *rowIndex), row.ServiceOffering)

		*rowIndex = *rowIndex + 1
	}
}

func createRow(row []string, dateInc int, timeSpent string) Row {
	date, err := time.Parse("01-02-06", row[0])
	if err != nil {
		fmt.Println(err, row[0])
		return Row{}
	}

	return Row{
		Date:                    date.AddDate(0, 0, dateInc).Format("2006-01-02"),
		Category:                row[1],
		User:                    row[2],
		Task:                    row[3],
		State:                   row[4],
		Time:                    timeSpent,
		Region:                  row[13],
		Country:                 row[14],
		Department:              row[15],
		OrganizationalHierarchy: row[16],
		FunctionalArea:          row[17],
		ServiceCatalog:          row[18],
		ServiceOffering:         row[19],
	}
}

func createRows(row []string) []Row {
	rows := make([]Row, 0)

	rows = append(rows, createRow(row, 0, row[5]))
	rows = append(rows, createRow(row, 1, row[6]))
	rows = append(rows, createRow(row, 2, row[7]))
	rows = append(rows, createRow(row, 3, row[8]))
	rows = append(rows, createRow(row, 4, row[9]))
	rows = append(rows, createRow(row, 5, row[10]))
	rows = append(rows, createRow(row, 6, row[11]))

	return rows
}

func main() {
	flag.Parse()

	if len(flag.Args()) != 0 {
		if flag.Args()[0] == "-h" || flag.Args()[0] == "--help" {
			flag.PrintDefaults()
			return
		}
	}

	inputFile, err := excelize.OpenFile(*input)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := inputFile.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	outputFile := excelize.NewFile()

	rows, err := inputFile.GetRows(*inputSheetName)
	if err != nil {
		fmt.Println(err)
		return
	}
	rows = rows[1:]

	createHeader(outputFile)

	var entries []Row
	for row := range rows {
		entries = append(entries, createRows(rows[row])...)
	}

	writeRows(outputFile, entries, new(int))
	outputFile.SaveAs(*output)
}
