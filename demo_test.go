package xlsx_test

import (
	"testing"

	"github.com/bingoohuang/xlsx"
)

func TestDemo1(t *testing.T) {
	x := xlsx.New()
	x.Write([]memberStat{
		{Total: 100, New: 50, Effective: 50},
		{Total: 200, New: 60, Effective: 140},
	})

	_ = x.Save("testdata/demo1.xlsx")
}

func TestDemo2(t *testing.T) {
	x := xlsx.New(xlsx.WithTemplate("testdata/template.xlsx"))
	x.Write([]memberStat{
		{Total: 100, New: 50, Effective: 50},
		{Total: 200, New: 60, Effective: 140},
	})

	_ = x.Save("testdata/demo2.xlsx")
}
