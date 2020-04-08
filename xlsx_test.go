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
	x, _ := xlsx.New(xlsx.WithTemplate("testdata/template.xlsx"))
	defer x.Close()

	writeData(t, time.Now(), x, "testdata/out_template.xlsx")
}

func Test1(t *testing.T) {
	now := startOfDay(time.Now())
	x, _ := xlsx.New()

	defer x.Close()

	writeData(t, now, x, "testdata/out_direct.xlsx")

	var memberStats []memberStat

	x2, _ := xlsx.New(xlsx.WithInputFile("testdata/out_direct.xlsx"))
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
	_ = x.Write([]memberStat{
		{Total: 100, New: 50, Effective: 50},
		{Total: 200, New: 60, Effective: 140},
	})

	_ = x.Write([]schedule{
		{Day: now, Num: 100, Subscribes: 500, PublicSubscribes: 400, PrivatesSubscribes: 100},
		{Day: now.AddDate(0, 0, -1), Num: 101, Subscribes: 501, PublicSubscribes: 401, PrivatesSubscribes: 101},
		{Day: now.AddDate(0, 0, -2), Num: 102, Subscribes: 502, PublicSubscribes: 402, PrivatesSubscribes: 102},
	})

	_ = x.Write([]orderStat{})

	assert.Nil(t, x.SaveToFile(file))
}

func startOfDay(t time.Time) time.Time {
	year, month, day := t.Date()
	return time.Date(year, month, day, 0, 0, 0, 0, t.Location())
}

type memberStat2 struct {
	xlsx.T `sheet:"会员"`

	Area      string `title:"区域" dataValidation:"Validation!A1:A3"`
	Total     int    `title:"会员总数"`
	New       int    `title:"其中：新增"`
	Effective int    `title:"其中：有效"`
}

func TestValidationTmpl(t *testing.T) {
	x, _ := xlsx.New(xlsx.WithTemplate("testdata/tmpl_validate.xlsx"))
	defer x.Close()

	err := x.Write([]memberStat2{
		{Area: "A", Total: 100, New: 50, Effective: 50},
		{Area: "B", Total: 200, New: 60, Effective: 140},
		{Area: "C", Total: 300, New: 70, Effective: 240},
	})

	assert.Nil(t, err)

	_ = x.SaveToFile("testdata/out_validation_tmpl.xlsx")
}

type memberStat22 struct {
	xlsx.T `sheet:"会员"`

	Area      string `title:"区域" dataValidation:"A22,B22,C22"`
	Total     int    `title:"会员总数"`
	New       int    `title:"其中：新增"`
	Effective int    `title:"其中：有效"`
}

func TestValidation(t *testing.T) {
	x, _ := xlsx.New()
	defer x.Close()

	err := x.Write([]memberStat22{
		{Area: "A22", Total: 100, New: 50, Effective: 50},
		{Area: "B22", Total: 200, New: 60, Effective: 140},
		{Area: "C22", Total: 300, New: 70, Effective: 240},
	})

	assert.Nil(t, err)

	_ = x.SaveToFile("testdata/out_validation.xlsx")
}

type memberStat23 struct {
	xlsx.T `sheet:"会员"`

	Area      string `title:"区域" dataValidation:"areas"`
	Total     int    `title:"会员总数"`
	New       int    `title:"其中：新增"`
	Effective int    `title:"其中：有效"`
}

func TestValidationWith(t *testing.T) {
	x, _ := xlsx.New(xlsx.WithValidations(map[string][]string{
		"areas": {"A23", "B23", "C23"},
	}))
	defer x.Close()

	err := x.Write([]memberStat23{
		{Area: "A23", Total: 100, New: 50, Effective: 50},
		{Area: "B23", Total: 200, New: 60, Effective: 140},
		{Area: "C23", Total: 300, New: 70, Effective: 240},
	})

	assert.Nil(t, err)

	_ = x.SaveToFile("testdata/out_validation_with.xlsx")
}

func TestParsePlaceholder(t *testing.T) {
	assert.Equal(t, xlsx.PlaceholderValue{
		PlaceholderVars:   []string{},
		PlaceholderQuotes: []string{},
		Content:           "Age",
		Parts:             []xlsx.PlaceholderPart{{Part: "Age"}},
	}, xlsx.ParsePlaceholder("Age"))

	assert.Equal(t, xlsx.PlaceholderValue{
		PlaceholderVars:   []string{"name"},
		PlaceholderQuotes: []string{"{{name}}"},
		Content:           "{{name}}",
		Parts:             []xlsx.PlaceholderPart{{Part: "{{name}}", Var: "name"}},
	}, xlsx.ParsePlaceholder("{{name}}"))

	assert.Equal(t, xlsx.PlaceholderValue{
		PlaceholderVars:   []string{"name", "age"},
		PlaceholderQuotes: []string{"{{name}}", "{{ age }}"},
		Content:           "{{name}} {{ age }}",
		Parts: []xlsx.PlaceholderPart{
			{Part: "{{name}}", Var: "name"},
			{Part: " ", Var: ""},
			{Part: "{{ age }}", Var: "age"},
		},
	}, xlsx.ParsePlaceholder("{{name}} {{ age }}"))

	assert.Equal(t, xlsx.PlaceholderValue{
		PlaceholderVars:   []string{"name", "age"},
		PlaceholderQuotes: []string{"{{name}}", "{{ age }}"},
		Content:           "Hello {{name}} world {{ age }}",
		Parts: []xlsx.PlaceholderPart{
			{Part: "Hello ", Var: ""},
			{Part: "{{name}}", Var: "name"},
			{Part: " world ", Var: ""},
			{Part: "{{ age }}", Var: "age"},
		},
	}, xlsx.ParsePlaceholder("Hello {{name}} world {{ age }}"))

	assert.Equal(t, xlsx.PlaceholderValue{
		PlaceholderVars:   []string{},
		PlaceholderQuotes: []string{},
		Content:           "Age{{",
		Parts: []xlsx.PlaceholderPart{
			{Part: "Age{{", Var: ""},
		},
	}, xlsx.ParsePlaceholder("Age{{"))

	plName := xlsx.ParsePlaceholder("{{name}}")
	vars, ok := plName.ParseVars("bingoohuang")
	assert.True(t, ok)
	assert.Equal(t, map[string]string{"name": "bingoohuang"}, vars)

	nameArgs := xlsx.ParsePlaceholder("{{name}} {{ age }}")
	vars, ok = nameArgs.ParseVars("bingoohuang 100")
	assert.True(t, ok)
	assert.Equal(t, map[string]string{"name": "bingoohuang", "age": "100"}, vars)

	nameArgs = xlsx.ParsePlaceholder("中国{{v1}}人民{{v2}}")
	vars, ok = nameArgs.ParseVars("中国中央人民政府")
	assert.True(t, ok)
	assert.Equal(t, map[string]string{"v1": "中央", "v2": "政府"}, vars)
}

type RegisterTable struct {
	ContactName  string    // 联系人
	Mobile       string    // 手机
	Landline     string    // 座机
	RegisterDate time.Time // 登记日期
	DeviceType   string    `placeholderCell:"C8"` // 类型
	Manufacturer string    // 生产厂家
	DeviceModern string    // 型号
}

func TestPlaceholder(t *testing.T) {
	x, _ := xlsx.New(xlsx.WithTemplatePlaceholder("testdata/placeholder.xlsx"))
	defer x.Close()

	now, _ := time.ParseInLocation("2006-01-02 15:04:05", "2020-04-08 20:53:11", time.Local)

	src := RegisterTable{
		ContactName:  "隔壁老王",
		Mobile:       "1234567890",
		Landline:     "010-1234567890",
		RegisterDate: now,
		DeviceType:   "A1",
		Manufacturer: "来弄你",
		DeviceModern: "X786",
	}
	err := x.Write(src)

	assert.Nil(t, err)

	_ = x.SaveToFile("testdata/out_placeholder.xlsx")

	x2, _ := xlsx.New(xlsx.WithTemplatePlaceholder("testdata/placeholder.xlsx"),
		xlsx.WithInputFile("testdata/out_placeholder.xlsx"))
	defer x2.Close()

	var v RegisterTable

	err = x2.Read(&v)

	assert.Nil(t, err)
	assert.Equal(t, src, v)
}
