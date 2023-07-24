package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"time"

	"github.com/bingoohuang/xlsx"
	"github.com/unidoc/unioffice/spreadsheet"
)

func main() {
	demo := flag.String("demo", "", "demo for what: placeholder")
	flag.Parse()

	switch *demo {
	case "1", "":
		demo1()
	case "placeholder", "ph":
		placeholder()
	default:
		readCellValues()
	}
}

func demo1() {
	type memberStat struct {
		Total     int `title:"会员总数" sheet:"会员"` // sheet可选，不声明则选择首个sheet页读写
		New       int `title:"其中：新增"`
		Effective int `title:"其中：有效"`
	}

	x, _ := xlsx.New()
	defer x.Close()

	x.Write([]memberStat{
		{Total: 100, New: 50, Effective: 50},
		{Total: 200, New: 60, Effective: 140},
	})
	x.SaveToFile("testdata/test1.xlsx")
}

func readCellValues() {
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

// RegisterTable 注册登记表信息.
type RegisterTable struct {
	ContactName  string    `asPlaceholder:"true"` // 联系人
	Mobile       string    // 手机
	Landline     string    // 座机
	RegisterDate time.Time `format:"yyyy-MM-dd"`  // 登记日期
	DeviceType   string    `placeholderCell:"C9"` // 类型
	Manufacturer string    // 生产厂家
	DeviceModern string    // 型号
}

func placeholder() {
	template := "testdata/placeholder.xlsx"
	x, _ := xlsx.New(xlsx.WithTemplate(template))
	defer x.Close()

	_ = x.Write(RegisterTable{
		ContactName:  "隔壁老王",
		Mobile:       "1234567890",
		Landline:     "010-1234567890",
		RegisterDate: time.Now(),
		DeviceType:   "A1",
		Manufacturer: "来弄你",
		DeviceModern: "X786",
	})

	out := "out_placeholder.xlsx"
	_ = x.SaveToFile(out)
	log.Printf("tempalte: %s, outout excel: %s", template, out)
}
