package main

import (
	"errors"
	"fmt"
	"log"
	"strings"

	// Vendor's libs
	"github.com/tealeg/xlsx"

	// Local libs
)

func main() {
	excelFileName := "test.xlsx"
	tbl1, err := ParseDoc(excelFileName)
	if err != nil {
		log.Printf("tbl1: %v", err)
	}
	fmt.Println(tbl1)
}

type Region struct {
	Origin struct{ X, Y int }
	Size   struct{ DX, DY int }
}

type Table struct {
	hdr  TableHeader
	data [][]string
}

type TableHeader struct {
	size TableSize
}

type TableSize struct {
	rows, cols int
}

func NewTable(rows, cols int) *Table {
	t := new(Table)

	// Set header
	t.hdr.size.rows = rows
	t.hdr.size.cols = cols

	// Create data structure
	t.data = make([][]string, rows)
	for i := range t.data {
		t.data[i] = make([]string, cols)
	}

	return t
}

func (t *Table) Header() []string {
	return t.data[0]
}

func (t *Table) Data() [][]string {
	return t.data[1:]
}

func (t *Table) SetDataAt(row, col int, data string) {
	t.data[row][col] = data
}

func (t *Table) String() string {
	res := "=== Header ===\n"
	res += strings.Join(t.Header(), "|")
	res += "\n---  Data  ---\n"

	data := t.Data()
	for x := range data {
		res += strings.Join(data[x], "|")
		res += "\n"
	}
	res += "===  End   ===\n"

	return res
}

func ParseDoc(path string) (*Table, error) {
	xlFile, err := xlsx.OpenFile(path)
	if err != nil {
		return nil, err
	}

	allRows := xlFile.Sheets[0].Rows

	reg, err := findTableRegeon(allRows)
	if err != nil {
		return nil, err
	}

	tbl := NewTable(reg.Size.DY, reg.Size.DX)
	for y, row := range allRows[reg.Origin.Y : reg.Origin.Y+reg.Size.DY] {
		for x, cell := range row.Cells[reg.Origin.X : reg.Origin.X+reg.Size.DX] {
			tbl.SetDataAt(y, x, getData(cell))
		}
	}

	return tbl, nil
}

func findTableRegeon(rows []*xlsx.Row) (*Region, error) {
	var reg *Region

	// Find origin and # of columns
	for y, row := range rows {
		for x, cell := range row.Cells {
			if hasBorder(cell) {
				if reg == nil {
					// set origin
					reg = new(Region)
					reg.Origin.X = x
					reg.Origin.Y = y
				}
				reg.Size.DX++
			} else {
				if reg != nil {
					break
				}
			}
		}
		if reg != nil {
			break
		}
	}

	// Find # of rows
	if reg != nil {
		for _, row := range rows[reg.Origin.Y:] {
			if hasBorder(row.Cells[reg.Origin.X]) {
				reg.Size.DY++
			} else {
				break
			}
		}

		return reg, nil
	} else {
		return nil, errors.New("No data")
	}
}

func getData(c *xlsx.Cell) string {
	// fmt.Printf("%s\n", cell.String())
	val, err := c.SafeFormattedValue()
	if err != nil {
		val = c.Value
	}

	return strings.Replace(val, `\ `, ` `, -1)
}

var (
	defaultBorder = *xlsx.DefaultBorder()
	emptyBorder   = xlsx.Border{}
)

func hasBorder(c *xlsx.Cell) bool {
	if c != nil {
		if b := c.GetStyle().Border; !(b == defaultBorder || b == emptyBorder) {
			return true
		}
	}
	return false
}

// vim: :ts=4:sw=4:noexpandtab:ai
