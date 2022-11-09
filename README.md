# gexcel

封装 github.com/qax-os/excelize 项目，方便业务使用

目前仅支持单 sheet，样式上仅支持：
* 设置 sheet 名称
* 设置通用列宽度
* 设置通用对齐方式

### 使用方式
`go get github.com/mao888/go-excel`
```go
sheetName := "测试sheet"
header := []string{"姓名", "性别", "年龄"}
content := [][]interface{}{
{"胡超", "男", 18},
}
file := NewExcel(sheetName, header, content)
//file.SetAlign(HeaderAlignCenter) //默认居中对齐
//file.SetColWidth(50) //默认宽度 20
data, err := file.Export() // 返回 data 类型为 *bytes.Buffer
```
