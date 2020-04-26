Go语言Excel读写库使用手册

1. 引入库 `import "github.com/bingoohuang/xlsx"`
1. 定义EXCEL模板，参见[示例](https://github.com/bingoohuang/xlsx/blob/master/testdata/template.xlsx)
1. 定义Go结构，与模板对应，例如:

    ```go
    type MemberStat struct {
        Total     int `title:"会员总数"`
        New       int `title:"其中：新增"`
        Effective int `title:"其中：有效"`
    }
    ```

1. 向Excel中写入，示例:

    ```go
    func main() {
        x, _ := xlsx.New(xlsx.WithTemplate("your-template.xlsx"))
        defer x.Close()
    
        x.Write([]MemberStat{
                {Total: 100, New: 50, Effective: 50},
                {Total: 200, New: 60, Effective: 140},
            })
        x.SaveToFile("testdata/test1.xlsx")
    }
    ```

1. 从Excel中读取，示例:

    ```go
    func main() {
        x, _ := xlsx.New(xlsx.WithExcel("testdata/test1.xlsx"))
        defer x.Close()
        
        var MemberStat []memberStat
        if err := x.Read(&memberStats); err != nil {
            fmt.Println(err)
        }
    }
    ```


1. 更多高级用法与示例，请[移步](https://github.com/bingoohuang/xlsx)
