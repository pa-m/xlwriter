package xlwriter

import (
	//"fmt"
	"os"
	//"testing"
	"time"
)

func ExampleXlWriter() {
	type T struct {
		B, C  []float64
		Dates []time.Time
	}
	y := T{B: []float64{1, 2, 3}, C: []float64{1, 2, 3}, Dates: []time.Time{time.Now()}}
	w := NewFile()
	w.SetSheetName(w.GetSheetName(0), "1st")
	w.WriteColumns("a", []int{1, 2, 3}, y)
	w.NewSheet("2nd")
	w.WriteColumns("a", []int{1, 2, 3})
	w.SaveAs("/tmp/test.xlsx")
	os.Remove("/tmp/test.xlsx")
	// Output:
}
