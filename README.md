# xlsx

golang mapping between xlsx and struct instances.

## Usage documentation

### Directly write excel file

```go
package main

import "github.com/bingoohuang/xlsx"

type memberStat struct {
	xlsx.T `sheet:"会员"`

	Total     int `title:"会员总数"`
	New       int `title:"其中：新增"`
	Effective int `title:"其中：有效"`
}

func main() {
	x := xlsx.New()
    x.Write([]memberStat{
    		{Total: 100, New: 50, Effective: 50},
    		{Total: 200, New: 60, Effective: 140},
    	})
    x.Save("testdata/test1.xlsx")
}
```

you will the the result excel file in [testdata/demo1.xlsx](testdata/demo1.xlsx) as follows:

![image](https://user-images.githubusercontent.com/1940588/77844342-a1d22580-71d8-11ea-8eb9-6f82f87c3a3a.png)

### Write excel with template file

```go
x := xlsx.New(xlsx.WithTemplate("testdata/template.xlsx"))
x.Write([]memberStat{
        {Total: 100, New: 50, Effective: 50},
        {Total: 200, New: 60, Effective: 140},
    })
x.Save("testdata/test1.xlsx")
```

you will the the result excel file in [testdata/demo2.xlsx](testdata/demo2.xlsx) as follows:

![image](https://user-images.githubusercontent.com/1940588/77844394-0ee5bb00-71d9-11ea-8671-6b36eb6a728b.png)

### Read excel with titled row

```go
var memberStats []memberStat

x := xlsx.New(xlsx.WithInputFile("testdata/test1.xlsx"))
if err := x.Read(&memberStats); err != nil {
    panic(err)
}

assert.Equal(t, []memberStat{
    {Total: 100, New: 50, Effective: 50},
    {Total: 200, New: 60, Effective: 140},
}, memberStats)
```
