package xlsx_test

import (
	"testing"
	"time"

	"github.com/bingoohuang/xlsx"

	"github.com/stretchr/testify/assert"
)

type memberStat struct {
	xlsx.T `sheet:"会员"`

	Total     int `title:"会员总数"`
	New       int `title:"其中：新增"`
	Effective int `title:"其中：有效"`
}

type schedule struct {
	xlsx.T `sheet:"排期"`

	Day                time.Time `title:"日期" format:"yyyy-MM-dd"`
	Num                int       `title:"排期数"`
	Subscribes         int       `title:"订课数"`
	PublicSubscribes   int       `title:"其中：小班课"`
	PrivatesSubscribes int       `title:"其中：私教课"`
}

type orderStat struct {
	xlsx.T `sheet:"订课情况"`

	Day   time.Time `title:"订单日期"`
	Time  int       `title:"人次"`
	Heads int       `title:"人数"`
}

func Test2(t *testing.T) {
	x := xlsx.New(xlsx.WithTemplate("testdata/template.xlsx"))

	writeData(t, x, "testdata/test2.xlsx")
}

func Test1(t *testing.T) {
	x := xlsx.New()
	writeData(t, x, "testdata/test1.xlsx")

	var memberStats []memberStat

	x = xlsx.New(xlsx.WithInputFile("testdata/test1.xlsx"))
	assert.Nil(t, x.Read(&memberStats))

	assert.Equal(t, []memberStat{
		{Total: 100, New: 50, Effective: 50},
		{Total: 200, New: 60, Effective: 140},
	}, memberStats)
}

func writeData(t *testing.T, x *xlsx.Xlsx, file string) {
	x.Write([]memberStat{
		{Total: 100, New: 50, Effective: 50},
		{Total: 200, New: 60, Effective: 140},
	})

	x.Write([]schedule{
		{Day: time.Now(), Num: 100, Subscribes: 500, PublicSubscribes: 400, PrivatesSubscribes: 100},
		{Day: time.Now().AddDate(0, 0, -1), Num: 101, Subscribes: 501, PublicSubscribes: 401, PrivatesSubscribes: 101},
		{Day: time.Now().AddDate(0, 0, -2), Num: 102, Subscribes: 502, PublicSubscribes: 402, PrivatesSubscribes: 102},
	})

	x.Write([]orderStat{})

	assert.Nil(t, x.Save(file))
}
