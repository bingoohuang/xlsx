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
