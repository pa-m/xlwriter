# pa-m/xlwriter

a go module to ease the use of the excellent "github.com/360EntSecGroup-Skylar/excelize" package

usage:

```go
import "github.com/pa-m/xlwriter"

func TestXlWriter() {
	w := NewFile()
	type T struct {
		B, C  []float64
		Dates []time.Time
	}
	y := T{B: []float64{1, 2, 3}, C: []float64{1, 2, 3}, Dates: []time.Time{time.Now()}}
	w.WriteColumns("a", []int{1, 2, 3}, y)
	w.SaveAs("/tmp/test.xlsx")

}
```
