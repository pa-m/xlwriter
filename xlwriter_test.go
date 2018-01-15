package xlwriter

import (
	"fmt"
	"os"
	"testing"
	"time"
)

func TestXlWriter(t *testing.T) {
	w := NewFile()
	type T struct {
		B, C  []float64
		Dates []time.Time
	}
	y := T{B: []float64{1, 2, 3}, C: []float64{1, 2, 3}, Dates: []time.Time{time.Now()}}
	w.WriteColumns("a", []int{1, 2, 3}, y)
	w.SaveAs("/tmp/test.xlsx")
	os.Remove("/tmp/test.xlsx")
	fmt.Println("TestXlWriter ok")
}

func ExampleXlWriter() {
	fmt.Println("hello")
	// Output: hello
}
