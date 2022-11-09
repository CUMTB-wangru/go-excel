package gexcel

import (
	"os"
	"testing"
)

func TestNewExcel(t *testing.T) {
	sheetName := "测试sheet"
	header := []string{"姓名", "性别", "年龄"}
	content := [][]interface{}{
		{"胡超", "男", 18},
	}
	file := NewExcel(sheetName, header, content)
	file.SetAlign(HeaderAlignCenter)
	file.SetColWidth(50)
	data, err := file.Export()
	if err != nil {
		t.Errorf("TestNewExcel err: %v ", err)
		return
	}
	err = os.WriteFile("testExcelExport.xlsx", data.Bytes(), 777)
	if err != nil {
		t.Errorf("TestNewExcel WriteFile err: %v ", err)
		return
	}
	t.Logf("TestNewExcel success")
}
