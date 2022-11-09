package gexcel

import (
	"bytes"
	"fmt"

	"github.com/xuri/excelize/v2"
)

type headerAlign string

const (
	HeaderAlignLeft   headerAlign = "left"
	HeaderAlignRight  headerAlign = "right"
	HeaderAlignCenter headerAlign = "center"
)

const (
	DefaultSheetName = "Sheet1"
	DefaultColWidth  = 20
)

type Excel struct {
	sheet   string
	header  []string
	content [][]interface{}

	align    headerAlign
	colWidth float64
}

// NewExcel 返回一个初始化过的对象. 目前只支持一个 sheet,可以设置 sheet name
func NewExcel(sheetName string, header []string, content [][]interface{}) *Excel {
	return &Excel{
		sheet:    sheetName,
		header:   header,
		content:  content,
		align:    HeaderAlignCenter,
		colWidth: DefaultColWidth,
	}
}

// NewEmptyExcel 返回一个默认的、空的 Excel 对象 目前只支持一个 sheet,可以设置 sheet name
func NewEmptyExcel() *Excel {
	return &Excel{
		sheet:    DefaultSheetName,
		header:   nil,
		content:  nil,
		align:    HeaderAlignCenter,
		colWidth: DefaultColWidth,
	}
}

// SetSheetName 设置 sheet 名称，默认为 Sheet1
func (e *Excel) SetSheetName(name string) {
	e.sheet = name
}

// SetHeader 设置 header
func (e *Excel) SetHeader(header []string) {
	e.header = header
}

// SetContent 设置 content
func (e *Excel) SetContent(content [][]interface{}) {
	e.content = content
}

// SetAlign 设置 对齐方式，默认居中对齐
func (e *Excel) SetAlign(align headerAlign) {
	e.align = align
}

// SetColWidth 设置 列宽度，默认 20
func (e *Excel) SetColWidth(width float64) {
	e.colWidth = width
}

func (e *Excel) Export() (*bytes.Buffer, error) {
	if e.header == nil || e.content == nil {
		return nil, fmt.Errorf("ExcelExpor err: header or content is nil")
	}
	if len(e.header) == 0 {
		return nil, fmt.Errorf("ExcelExpor err: header is empty slice")
	}
	f := excelize.NewFile()
	defer f.Close()
	sheetName := e.sheet
	if sheetName != DefaultSheetName {
		sheetName = f.GetSheetName(f.NewSheet(e.sheet))
		f.DeleteSheet(DefaultSheetName)
	}
	style, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: string(e.align),
			Vertical:   "center",
			WrapText:   true,
		},
	})
	if err != nil {
		return nil, fmt.Errorf("ExcelExport create style err: %v", err)
	}

	firstColName := getColumnName(0)
	lastColName := getColumnName(len(e.header) - 1)

	//设置表格宽度
	if err = f.SetColWidth(sheetName, string(firstColName), string(lastColName), e.colWidth); err != nil {
		return nil, fmt.Errorf("ExcelExport set width err: %v", err)
	}
	//设置样式
	if err = f.SetColStyle(sheetName, fmt.Sprintf("%s:%s", string(firstColName), string(lastColName)), style); err != nil {
		return nil, fmt.Errorf("ExcelExport set style err: %v", err)
	}

	//生成header
	for colIndex, colValue := range e.header {
		colName := getColumnName(colIndex)
		axis := fmt.Sprintf("%s%d", colName, 1)
		if err = f.SetCellValue(sheetName, axis, colValue); err != nil {
			return nil, fmt.Errorf("ExcelExport set cell value err: %v", err)
		}
	}

	//生成数据
	rowBase := 2
	for row, rowData := range e.content {
		adjRow := rowBase + row
		if len(rowData) != len(e.header) {
			return nil, fmt.Errorf("ExcelExport err: row len not equal header len")
		}
		for colIndex, _ := range e.header {
			celData := rowData[colIndex]
			colName := getColumnName(colIndex)
			axis := fmt.Sprintf("%s%d", colName, adjRow)
			if err = f.SetCellValue(sheetName, axis, celData); err != nil {
				return nil, fmt.Errorf("ExcelExport set cell value err: %v", err)
			}
		}
	}
	return f.WriteToBuffer()
}

// maxCharCount 最多26个字符A-Z
const maxCharCount = 26

func getColumnName(columnIdx int) []byte {
	const A = 'A'
	if columnIdx < maxCharCount {
		slice := make([]byte, 0, 1)
		return append(slice, byte(A+columnIdx))
	} else {
		return append(getColumnName(columnIdx/maxCharCount-1), byte(A+columnIdx%maxCharCount))
	}
}
