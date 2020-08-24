package main

import (
	"fmt"
	"os"

	"github.com/szyhf/go-excel"
)

type DataImport struct {
	Key   string `xlsx:"column(key1#key2#key3)"`
	Value string `xlsx:"column(value)"`
}

func main() {
	// goexcel ../../testdata/埋点导入模板-pdf.xlsx
	if len(os.Args) < 1 {
		fmt.Fprintln(os.Stderr, "Usage: goexcel demo_imnput.xlsx")
		os.Exit(0)
	}

	// will assume the sheet name as "Standard" from the struct name.
	var dataImports []DataImport
	if err := excel.UnmarshalXLSX(os.Args[1], &dataImports); err != nil {
		panic(err)
	}

	for i, v := range dataImports {
		fmt.Printf("%d: %+v\n", i+1, v)
	}
}
