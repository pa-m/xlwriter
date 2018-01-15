package xlwriter

import (
	//   "errors"
	"fmt"
	"reflect"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
)

type XlWriterStyles struct{ Default, Bold int }

// XlWriter wraps excelize with a context and provide easy methods like WriteColumns,WriteStruct
type XlWriter struct {
	*excelize.File
	Sheet       string
	L, C, Style int
	Styles      XlWriterStyles
}

func NewFile() XlWriter {
	xlsx := excelize.NewFile()
	return XlWriter{xlsx, "Sheet1", 0, 0, 0, DefaultStyles(xlsx)}
}
func DefaultStyles(File *excelize.File) XlWriterStyles {

	Bold, err := File.NewStyle(`{"font":{"bold":true}}`)
	if err != nil {
		panic(err)
	}
	return XlWriterStyles{Bold: Bold}
}

func (w *XlWriter) Ref(l, c int) string {
	cell := fmt.Sprintf("%c%d", 'A'+(c%26), l+1)
	if c >= 26 {
		c0 := c/26 - 1
		cell = fmt.Sprintf("%c%s", 'A'+c0, cell)
	}
	//fmt.Println(w.Sheet,cell)
	return cell
}

func (w *XlWriter) WriteString(v string) {
	cell := w.Ref(w.L, w.C)
	if w.Style > 0 {
		w.File.SetCellStyle(w.Sheet, cell, cell, w.Style)
	}
	w.File.SetCellValue(w.Sheet, cell, v)
}

func (w *XlWriter) WriteValue(i interface{}) {
	cell := w.Ref(w.L, w.C)
	v, ok := i.(reflect.Value)
	if !ok {
		v = reflect.ValueOf(i)
	}
	switch v.Type().String() {
	case "float64":
		w.File.SetCellValue(w.Sheet, cell, v.Float())
	case "int":
		w.File.SetCellValue(w.Sheet, cell, v.Int())
	case "string":
		w.WriteString(v.String())
	case "time.Time":
		w.File.SetCellValue(w.Sheet, cell, v.Interface().(time.Time).String()[:16])
	default:
		panic(fmt.Sprintf("ExcelWriteValue %s", v.Type().String()))
	}
}

// WriteColumns takes i..n name,slices or structs having slice members

func (w *XlWriter) WriteColumns(its ...interface{}) {
	for _, it := range its {
		v, ok := it.(reflect.Value)
		if !ok {
			v = reflect.ValueOf(it)
		}
		isStructWithSlices := false
		if v.Kind() == reflect.Struct {
			for i := 0; i < v.NumField(); i++ {
				if v.Field(i).Kind() == reflect.Slice {
					isStructWithSlices = true
				}
			}
		}
		if isStructWithSlices {
			//fmt.Println("sws")
			for i := 0; i < v.NumField(); i++ {
				if v.Field(i).Kind() == reflect.Slice {
					//fmt.Println("-", v.Type().Field(i).Name)
					w.WriteColumns(v.Type().Field(i).Name, v.Field(i))
				}
			}
		} else if v.Kind() == reflect.Slice {
			//fmt.Println("slice", v.Len())
			for j := 0; j < v.Len(); j++ {
				w.WriteValue(v.Index(j))
				w.L++
				//fmt.Printf("%s %g ",fmt.Sprintf("%c%d",'A'+i,j+2),values.Index(j).Float())
			}
			w.C++
			w.L = 0
		} else if v.Kind() == reflect.String {
			//fmt.Println("string", v.String())
			w.Style = w.Styles.Bold
			w.WriteString(v.String())
			w.L++
			w.Style = 0
		}
	}
}

// WriteStruct takes 1..n title and a struct. it write the struct vertically on 2 columns

func (w *XlWriter) WriteStructs(its ...interface{}) {
	for _, it := range its {
		v, ok := it.(reflect.Value)
		if !ok {
			v = reflect.ValueOf(it)
		}
		if v.Kind() == reflect.String {
			w.Style = w.Styles.Bold
			w.WriteString(v.String())
			w.L++
			w.Style = 0
		} else if v.Kind() == reflect.Struct {
			vT := v.Type()
			for i := 0; i < v.NumField(); i++ {
				w.WriteString(vT.Field(i).Name)
				w.C++
				w.WriteValue(v.Field(i))
				w.C--
				w.L++
			}
			w.L++
		}
	}
}

func (w *XlWriter) NewSheet(newName string) int {
	r := w.File.NewSheet(newName)
	w.Sheet = newName
	return r
}

func (w *XlWriter) SetSheetName(oldName, newName string) {
	w.File.SetSheetName(oldName, newName)
	if w.Sheet == oldName {
		w.Sheet = newName
	}
}

func (w *XlWriter) SaveAs(filename string) { w.File.SaveAs(filename) }
