package xlsx_test

import (
	"testing"

	"github.com/bingoohuang/xlsx"
)

func TestDemo1(t *testing.T) {
	x, _ := xlsx.New()
	defer x.Close()

	_ = x.Write([]memberStat{
		{Total: 100, New: 50, Effective: 50},
		{Total: 200, New: 60, Effective: 140},
	})

	_ = x.SaveToFile("testdata/out_demo1.xlsx")
}

func TestDemo2(t *testing.T) {
	x, _ := xlsx.New(xlsx.WithTemplateFile("testdata/template.xlsx"))
	defer x.Close()

	_ = x.Write([]memberStat{
		{Total: 100, New: 50, Effective: 50},
		{Total: 200, New: 60, Effective: 140},
	})

	_ = x.SaveToFile("testdata/out_demo2.xlsx")
}
