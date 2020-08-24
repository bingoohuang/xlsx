package main

import (
	"fmt"
	"log"
	"os"

	"github.com/bingoohuang/xlsx"
	"github.com/unidoc/unioffice/spreadsheet"
)

func main() {
	fn := ""

	if len(os.Args) <= 1 {
		fmt.Fprintf(os.Stderr, "Usage: xlsx demo.xlsx")
		os.Exit(0)
	}

	fn = os.Args[1]

	wb, err := spreadsheet.Open(fn)
	if err != nil {
		log.Fatalf("error opening reference sheet: %s", err)
	}

	formulaCount := 0

	for i, sheet := range wb.Sheets() {
		log.Printf("Sheet at %d:%s\n", i+1, sheet.Name())

		for j, row := range sheet.Rows() {
			log.Printf("Row at %d\n", j+1)

			for k, cell := range row.Cells() {
				s := xlsx.GetCellString(cell)
				log.Printf("Cell[%d:%d] %s\n", j+1, k+1, s)
			}
		}
	}

	log.Printf("evaluated %d formulas from %s sheet", formulaCount, fn)
}
