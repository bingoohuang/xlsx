package main

import (
	"time"

	"github.com/bingoohuang/xlsx"
	"github.com/spf13/pflag"
)

// RegisterTable 注册登记表信息
type RegisterTable struct {
	xlsx.T `asPlaceholder:"true"`

	ContactName  string    // 联系人
	Mobile       string    // 手机
	Landline     string    // 座机
	RegisterDate time.Time `format:"yyyy-MM-dd"`  // 登记日期
	DeviceType   string    `placeholderCell:"C9"` // 类型
	Manufacturer string    // 生产厂家
	DeviceModern string    // 型号
}

func main() {
	x, _ := xlsx.New(xlsx.WithTemplate("testdata/placeholder.xlsx"))
	defer x.Close()

	pflag.String("logrus", "", "logrus")
	pflag.Parse()

	_ = x.Write(RegisterTable{
		ContactName:  "隔壁老王",
		Mobile:       "1234567890",
		Landline:     "010-1234567890",
		RegisterDate: time.Now(),
		DeviceType:   "A1",
		Manufacturer: "来弄你",
		DeviceModern: "X786",
	})

	_ = x.SaveToFile("out_placeholder.xlsx")
}
