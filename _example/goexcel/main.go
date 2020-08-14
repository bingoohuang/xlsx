package main

import (
	"fmt"
	"github.com/szyhf/go-excel"
	"os"
)

type DataImport struct {
	Key   string `xlsx:"column(key1#key2#key3)"`
	Value string `xlsx:"column(value)"`
}

func main() {
	// goexcel ../../testdata/埋点导入模板-pdf.xlsx

	// will assume the sheet name as "Standard" from the struct name.
	var dataImports []DataImport
	if err := excel.UnmarshalXLSX(os.Args[1], &dataImports); err != nil {
		panic(err)
	}

	for i, v := range dataImports {
		fmt.Printf("%d: %+v\n", i+1, v)
	}
}
