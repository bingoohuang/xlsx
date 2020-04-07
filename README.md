# xlsx

golang mapping between xlsx and struct instances.

本库提供高层golang的struct切片和excel文件的映射，避免直接处理底层sheet/row/cells等细节。

本库底层使用[unioffice](https://github.com/unidoc/unioffice)，其提供了比[360EntSecGroup-Skylar/excelize](https://github.com/bingoohuang/xlsx)更加友好的API。

还有另外一个比较活跃的底层实现[tealeg/xlsx](https://github.com/tealeg/xlsx)尚未认证。

## Usage documentation

### Directly write excel file

```go
package main

import "github.com/bingoohuang/xlsx"

type memberStat struct {
	xlsx.T `sheet:"会员"` // 可选，如果不声明，会默认选择第一个sheet页进行读写

	Total     int `title:"会员总数"`
	New       int `title:"其中：新增"`
	Effective int `title:"其中：有效"`
}

func main() {
	x, _ := xlsx.New()
    defer x.Close()

    x.Write([]memberStat{
    		{Total: 100, New: 50, Effective: 50},
    		{Total: 200, New: 60, Effective: 140},
    	})
    x.SaveToFile("testdata/test1.xlsx")
}
```

you will get the result excel file in [testdata/demo1.xlsx](testdata/demo1.xlsx) as the following:

![image](https://user-images.githubusercontent.com/1940588/77844342-a1d22580-71d8-11ea-8eb9-6f82f87c3a3a.png)

### Write excel with template file

```go
x, _ := xlsx.New(xlsx.WithTemplate("testdata/template.xlsx"))
defer x.Close()

x.Write([]memberStat{
        {Total: 100, New: 50, Effective: 50},
        {Total: 200, New: 60, Effective: 140},
    })
x.SaveToFile("testdata/test1.xlsx")
```

you will get the result excel file in [testdata/demo2.xlsx](testdata/demo2.xlsx) as the following:

![image](https://user-images.githubusercontent.com/1940588/77844394-0ee5bb00-71d9-11ea-8671-6b36eb6a728b.png)

### Read excel with titled row

```go
var memberStats []memberStat

x, _ := xlsx.New(xlsx.WithInputFile("testdata/test1.xlsx"))
defer x.Close()

if err := x.Read(&memberStats); err != nil {
    panic(err)
}

assert.Equal(t, []memberStat{
    {Total: 100, New: 50, Effective: 50},
    {Total: 200, New: 60, Effective: 140},
}, memberStats)
```


### create data validation

1. Method 1: use template sheet to list the validation datas like:

![image](https://user-images.githubusercontent.com/1940588/78579374-692eed80-7863-11ea-931e-ab74035baa1b.png)

then declare the data validation tag `dataValidation` like:

```go
type Member struct {
	Area      string `title:"区域" dataValidation:"Validation!A1:A3"`
	Total     int    `title:"会员总数"`
	New       int    `title:"其中：新增"`
	Effective int    `title:"其中：有效"`
}
```

2. Method 2: directly give the list in tag `dataValidation` with comma-separated like:

```go
type Member struct {
	Area      string `title:"区域" dataValidation:"A,B,C"`
	Total     int    `title:"会员总数"`
	New       int    `title:"其中：新增"`
	Effective int    `title:"其中：有效"`
}
```

3. Method 3: programmatically declares and specified the key name in the tag `dataValidation` like:

```go
type Member struct {
	Area      string `title:"区域" dataValidation:"areas"`
	Total     int    `title:"会员总数"`
	New       int    `title:"其中：新增"`
	Effective int    `title:"其中：有效"`
}

func demo() {
	x, _ := xlsx.New(xlsx.WithValidations(map[string][]string{
		"areas": {"A23", "B23", "C23"},
	}))
	defer x.Close()

	_ = x.Write([]memberStat23{
		{Area: "A23", Total: 100, New: 50, Effective: 50},
		{Area: "B23", Total: 200, New: 60, Effective: 140},
		{Area: "C23", Total: 300, New: 70, Effective: 240},
	})

	_ = x.SaveToFile("result.xlsx")
}
```

### 占位模板

![image](https://user-images.githubusercontent.com/1940588/78628536-f9eae500-78c6-11ea-90f0-29b5bb3a4610.png)

```go
type RegisterTable struct {
	ContactName  string    // 联系人
	Mobile       string    // 手机
	Landline     string    // 座机
	RegisterDate time.Time // 登记日期
	DeviceType   string    `placeholderCell:"C9"` // 类型
	Manufacturer string    // 生产厂家
	DeviceModern string    // 型号
}

func demo() {
	x, _ := xlsx.New(xlsx.WithTemplatePlaceholder("testdata/placeholder.xlsx"))
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


	_ = x.SaveToFile("testdata/out_placeholder.xlsx")
}

```

![image](https://user-images.githubusercontent.com/1940588/78628579-17b84a00-78c7-11ea-84bc-1a7e192ee06c.png)