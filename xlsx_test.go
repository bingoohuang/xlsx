package xlsx_test

import (
	"encoding/json"
	"errors"
	"fmt"
	"io/ioutil"
	"os"
	"testing"
	"time"

	"github.com/bingoohuang/xlsx"

	"github.com/stretchr/testify/assert"
)

// ReadBytes reads bytes from the file.
func ReadBytes(filename string) []byte {
	b, _ := ioutil.ReadFile(filename)
	return b
}

func ExampleXlsx() {
	type (
		HostInfo struct {
			ServerName         string    `title:"主机名称" json:"serverName"`
			ServerHostname     string    `title:"主机hostname" json:"serverHostname"`
			ServerIP           string    `title:"主机IP" json:"serverIp"`
			ServerUserRtx      string    `json:"serverUserRtx"`
			Status             string    `json:"status"` // 状态：0正常 1删除
			InstanceID         string    `title:"实例ID" json:"instanceId"`
			Region             string    `title:"服务器可用区" json:"region"`
			CreateTime         time.Time `json:"createTime"` // 创建时间
			UpdateTime         time.Time `json:"updateTime"` // 修改时间
			ServerUserFullName string    `title:"主机负责人(rtx)" json:"serverUserFullName"`
		}
		Rsp struct {
			Status  int        `json:"status"`
			Message string     `json:"message"`
			Data    []HostInfo `json:"data"`
		}
	)

	var r Rsp

	err := json.Unmarshal(ReadBytes("testdata/hostinfos.json"), &r)
	fmt.Println("Unmarshal", err == nil)

	x, _ := xlsx.New(xlsx.WithTemplate("testdata/hostinfos_template.xlsx"))
	defer x.Close()

	err = x.Write(r.Data, xlsx.WithSheetName("FirstSheet"))
	fmt.Println("Write", err == nil)

	r.Data[0].ServerName += "第2页啦"
	_ = x.Write(r.Data, xlsx.WithSheetName("SecondSheet"))

	err = x.SaveToFile("testdata/out_hostinfos.xlsx")
	fmt.Println("SaveToFile", err == nil)
	// Output:
	// Unmarshal true
	// Write true
	// SaveToFile true
}

func ExampleNew() {
	x, _ := xlsx.New()
	defer x.Close()

	_ = x.Write([]memberStat{
		{Total: 100, New: 50, Effective: 50},
		{Total: 200, New: 60, Effective: 140},
	})

	err := x.SaveToFile("testdata/out_demo1.xlsx")

	// See: https://golang.org/pkg/testing/#hdr-Examples
	fmt.Println("Write", err == nil)
	// Output: Write true
}

func TestMerge(t *testing.T) {
	x, _ := xlsx.New()
	defer x.Close()

	_ = x.Write([]memberStat{
		{Total: 100, New: 50, Effective: 50},
		{Total: 100, New: 60, Effective: 50},
		{Total: 200, New: 60, Effective: 140},
		{Total: 200, New: 60, Effective: 140},
	}, xlsx.WithMergeColsMode(xlsx.MergeCols))

	err := x.SaveToFile("testdata/out_demo1_merge.xlsx")
	assert.Nil(t, err)
}

func TestMergeAlign(t *testing.T) {
	x, _ := xlsx.New()
	defer x.Close()

	_ = x.Write([]memberStat{
		{Total: 100, New: 50, Effective: 50},
		{Total: 100, New: 60, Effective: 50},
		{Total: 200, New: 60, Effective: 140},
		{Total: 200, New: 60, Effective: 140},
	}, xlsx.WithMergeColsMode(xlsx.MergeColsAlign))

	err := x.SaveToFile("testdata/out_demo1_merge_align.xlsx")
	assert.Nil(t, err)
}

func TestWithTemplateMerge(t *testing.T) {
	x, _ := xlsx.New(xlsx.WithTemplate("testdata/template.xlsx"))
	defer x.Close()

	_ = x.Write([]memberStat{
		{Total: 100, New: 50, Effective: 50},
		{Total: 100, New: 60, Effective: 50},
		{Total: 200, New: 60, Effective: 140},
		{Total: 200, New: 60, Effective: 140},
	}, xlsx.WithMergeColsMode(xlsx.MergeCols))

	err := x.SaveToFile("testdata/out_demo2_merge.xlsx")
	assert.Nil(t, err)
}

func TestWithTemplateMergeColsAlign(t *testing.T) {
	x, _ := xlsx.New(xlsx.WithTemplate("testdata/template.xlsx"))
	defer x.Close()

	_ = x.Write([]memberStat{
		{Total: 100, New: 50, Effective: 50},
		{Total: 100, New: 60, Effective: 50},
		{Total: 200, New: 60, Effective: 140},
		{Total: 200, New: 60, Effective: 140},
	}, xlsx.WithMergeColsMode(xlsx.MergeColsAlign))

	err := x.SaveToFile("testdata/out_demo2_merge_align.xlsx")
	assert.Nil(t, err)
}

func ExampleWithTemplate() {
	x, _ := xlsx.New(xlsx.WithTemplate("testdata/template.xlsx"))
	defer x.Close()

	_ = x.Write([]memberStat{
		{Total: 100, New: 50, Effective: 50},
		{Total: 200, New: 60, Effective: 140},
	})

	err := x.SaveToFile("testdata/out_demo2.xlsx")
	fmt.Println("Write", err == nil)
	// Output: Write true
}

type memberStat struct {
	Total     int `title:"会员总数" sheet:"会员"`
	New       int `title:"其中：新增"`
	Effective int `title:"其中：有效"`
}

type schedule struct {
	Day                time.Time `title:"日期" format:"yyyy-MM-dd" sheet:"排期"`
	Num                int       `title:"排期数"`
	Subscribes         int       `title:"订课数"`
	PublicSubscribes   int       `title:"其中：小班课"`
	PrivatesSubscribes int       `title:"其中：私教课"`
}

type orderStat struct {
	Day   time.Time `title:"订单日期" format:"yyyy-MM-dd" sheet:"订课情况"`
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

	x2, _ := xlsx.New(xlsx.WithExcel("testdata/out_direct.xlsx"))
	defer x2.Close()

	assert.Nil(t, x2.Read(&memberStats))

	assert.Equal(t, []memberStat{
		{Total: 100, New: 50, Effective: 50},
		{Total: 200, New: 60, Effective: 140},
		{Total: 300, New: 70, Effective: 150},
		{Total: 400, New: 80, Effective: 160},
		{Total: 500, New: 90, Effective: 180},
		{Total: 600, New: 96, Effective: 186},
		{Total: 700, New: 97, Effective: 187},
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
		{Total: 300, New: 70, Effective: 150},
		{Total: 400, New: 80, Effective: 160},
		{Total: 500, New: 90, Effective: 180},
		{Total: 600, New: 96, Effective: 186},
		{Total: 700, New: 97, Effective: 187},
	})

	_ = x.Write([]schedule{
		{Day: now, Num: 100, Subscribes: 500, PublicSubscribes: 400, PrivatesSubscribes: 100},
		{Day: now.AddDate(0, 0, -1), Num: 101, Subscribes: 501, PublicSubscribes: 401, PrivatesSubscribes: 101},
		{Day: now.AddDate(0, 0, -2), Num: 102, Subscribes: 502, PublicSubscribes: 402, PrivatesSubscribes: 102},
	})

	_ = x.Write(orderStat{
		Day:   time.Now(),
		Time:  10,
		Heads: 20,
	})

	assert.Nil(t, x.SaveToFile(file))
}

func startOfDay(t time.Time) time.Time {
	year, month, day := t.Date()
	return time.Date(year, month, day, 0, 0, 0, 0, t.Location())
}

type memberStat2 struct {
	Area      string `title:"区域" dataValidation:"Validation!A1:A3" sheet:"会员"`
	Total     int    `title:"=会员总数"`
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
	Area      string `title:"区域" dataValidation:"A22,B22,C22" sheet:"会员"`
	Total     int    `title:"=会员总数"`
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

	assert.Nil(t, x.Write([]memberStat22{}))

	_ = x.SaveToFile("testdata/out_validation.xlsx")
}

type memberStat23 struct {
	Area      string `title:"区域" dataValidation:"areas" sheet:"会员"`
	Total     int    `title:"=会员总数"`
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

type RegisterTable struct {
	ContactName  string    `asPlaceholder:"true"` // 联系人
	Mobile       string    // 手机
	Landline     string    // 座机
	RegisterDate time.Time `format:"yyyy-MM-dd"`  // 登记日期
	DeviceType   string    `placeholderCell:"C8"` // 类型
	Manufacturer string    // 生产厂家
	DeviceModern string    // 型号
}

func TestPlaceholder(t *testing.T) {
	bs, _ := ioutil.ReadFile("testdata/placeholder.xlsx")
	x, _ := xlsx.New(xlsx.WithTemplate(bs))

	defer x.Close()

	now, _ := time.ParseInLocation("2006-01-02", "2020-04-08", time.Local)

	src := RegisterTable{
		ContactName:  "隔壁老王",
		Mobile:       "1234567890",
		Landline:     "010-1234567890",
		RegisterDate: now,
		DeviceType:   "A1",
		Manufacturer: "来弄你",
		DeviceModern: "X786",
	}

	assert.Nil(t, x.Write(&src))
	assert.Nil(t, x.Write(src))
	assert.Nil(t, x.Write([]RegisterTable{src}))

	_ = x.SaveToFile("testdata/out_placeholder.xlsx")

	file, _ := os.Open("testdata/placeholder.xlsx")
	defer file.Close()
	x2, _ := xlsx.New(
		xlsx.WithTemplate(file),
		xlsx.WithExcel("testdata/out_placeholder.xlsx"))

	defer x2.Close()

	var v RegisterTable

	assert.Nil(t, x2.Read(&v))
	assert.Equal(t, src, v)
	assert.NotNil(t, x2.Read(v))
}

func TestIgnoreEmptyRows(t *testing.T) {
	type sch1 struct {
		Day                time.Time `title:"日期" format:"yyyy-MM-dd" sheet:"排期"`
		Num                int       `title:"排期数"`
		Subscribes         int       `title:"订课数"`
		PublicSubscribes   int       `title:"其中：小班课"`
		PrivatesSubscribes int       `title:"其中：私教课"`
	}

	x, err := xlsx.New(xlsx.WithExcel("testdata/template.xlsx"))
	assert.Nil(t, err)

	var schs1 []sch1

	err = x.Read(&schs1)
	assert.Nil(t, err)
	assert.Equal(t, 0, len(schs1))

	type sch2 struct {
		Day                time.Time `title:"日期" format:"yyyy-MM-dd" sheet:"排期" ignoreEmptyRows:"false"`
		Num                int       `title:"排期数"`
		Subscribes         int       `title:"订课数"`
		PublicSubscribes   int       `title:"其中：小班课"`
		PrivatesSubscribes int       `title:"其中：私教课"`
	}

	var schs2 []sch2

	err = x.Read(&schs2)
	assert.Nil(t, err)
	assert.Equal(t, 102, len(schs2))
}

type DataImport struct {
	Key   string `title:"key"`
	Value string `title:"value"`
}

func TestWangMengPdf(t *testing.T) {
	x, err := xlsx.New(xlsx.WithExcel("testdata/埋点导入模板-pdf.xlsx"))
	assert.Nil(t, err)

	var dataImports []DataImport
	err = x.Read(&dataImports)
	assert.Nil(t, err)

	for i, v := range dataImports {
		fmt.Printf("%d: %+v\n", i+1, v)
	}
}

func TestMaiDianDaoRu(t *testing.T) {
	x, err := xlsx.New(xlsx.WithExcel("testdata/埋点导入模板-813.xlsx"))
	assert.Nil(t, err)

	var dataImports []DataImport

	err = x.Read(&dataImports)
	assert.Nil(t, err)
	assert.Equal(t, []DataImport{
		{Key: "签名验证业务#PKCS7验签(带原文)#verifySignedDataByP7Attach", Value: "数据验签带原文"},
		{Key: "签名验证业务#PKCS7签名(不带原文)#signDataByP7Attach", Value: "数据签名不带原文"},
		{Key: "签名验证业务#PKCS7验签(不带原文)#verifySignedDataByP7Detach", Value: "数据验签不带原文"},
		{Key: "签名验证业务#PKCS7签名(不带原文)#signHashedDataPkcs7_detach", Value: "Hash数据签名不带原文"},
		{Key: "签名验证业务#PKCS7验签(不带原文)#signHashedDataPkcs7_detach", Value: "Hash数据验签不带原文"},
	}, dataImports)
}

func TestLocationError(t *testing.T) {
	x, err := xlsx.New(xlsx.WithExcel("testdata/bad.xlsx"))
	assert.Nil(t, err)

	var dataImports []DataImport
	err = x.Read(&dataImports)
	assert.True(t, errors.Is(err, xlsx.ErrFailToLocationTitleRow))
}
