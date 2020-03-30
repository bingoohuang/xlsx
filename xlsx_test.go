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
	defer x.Close()

	writeData(t, time.Now(), x, "testdata/test2.xlsx")
}

func Test1(t *testing.T) {
	now := startOfDay(time.Now())
	x := xlsx.New()

	defer x.Close()

	writeData(t, now, x, "testdata/test1.xlsx")

	var memberStats []memberStat

	x2 := xlsx.New(xlsx.WithInputFile("testdata/test1.xlsx"))
	defer x2.Close()

	assert.Nil(t, x2.Read(&memberStats))

	assert.Equal(t, []memberStat{
		{Total: 100, New: 50, Effective: 50},
		{Total: 200, New: 60, Effective: 140},
	}, memberStats)

	var schedules []schedule

	assert.Nil(t, x2.Read(&schedules))

	assert.Equal(t, []schedule{
		{Day: now, Num: 100, Subscribes: 500, PublicSubscribes: 400, PrivatesSubscribes: 100},
		{Day: now.AddDate(0, 0, -1), Num: 101, Subscribes: 501, PublicSubscribes: 401, PrivatesSubscribes: 101},
		{Day: now.AddDate(0, 0, -2), Num: 102, Subscribes: 502, PublicSubscribes: 402, PrivatesSubscribes: 102},
	}, schedules)
}

func writeData(t *testing.T, now time.Time, x *xlsx.Xlsx, file string) {
	x.Write([]memberStat{
		{Total: 100, New: 50, Effective: 50},
		{Total: 200, New: 60, Effective: 140},
	})

	x.Write([]schedule{
		{Day: now, Num: 100, Subscribes: 500, PublicSubscribes: 400, PrivatesSubscribes: 100},
		{Day: now.AddDate(0, 0, -1), Num: 101, Subscribes: 501, PublicSubscribes: 401, PrivatesSubscribes: 101},
		{Day: now.AddDate(0, 0, -2), Num: 102, Subscribes: 502, PublicSubscribes: 402, PrivatesSubscribes: 102},
	})

	x.Write([]orderStat{})

	assert.Nil(t, x.SaveToFile(file))
}

func startOfDay(t time.Time) time.Time {
	year, month, day := t.Date()
	return time.Date(year, month, day, 0, 0, 0, 0, t.Location())
}
