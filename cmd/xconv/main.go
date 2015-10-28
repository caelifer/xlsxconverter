package main

import (
	"errors"
	"fmt"

	// Vendor's libs
	"github.com/tealeg/xlsx"

	// Local libs
)

func main() {
	excelFileName := "test.xlsx"
	tbl, err := ParseDoc(excelFileName)
	fmt.Println(tbl, err)
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

func (t *Table) CellAt(row, col int) string {
	return t.data[row+1][col]
}

func ParseDoc(path string) (*Table, error) {
	xlFile, err := xlsx.OpenFile(path)
	if err != nil {
		return nil, err
	}

	getData := func(c *xlsx.Cell) string {
		// fmt.Printf("%s\n", cell.String())
		val, err := c.SafeFormattedValue()
		if err != nil {
			val = c.Value
		}
		return val
	}

	mtrx := make([][]string, 0, 10) // 10 rows by default
	sheet := xlFile.Sheets[0]       // Always fist sheet

	for nr, row := range sheet.Rows {
		if len(row.Cells) == 0 {
			continue
		}

		// log.Printf("===\nbdr(%#v) == defBdr(%#v) = %v\nbdr(%#v) == emptyBdr(%#v) = %v", bdr, defBdr, bdr == defBdr, bdr, emptyBdr, bdr == emptyBdr)

		// Find first cell with border set
		offset := 0
		hbFlag := false
		for i, c := range row.Cells {
			if hasBorder(c) {
				offset = i
				hbFlag = true
				break
			}
		}

		if !hbFlag {
			continue
		}

		fmt.Printf("%03d ", nr)

		r := make([]string, 0, len(row.Cells)-offset)
		for nc, cell := range row.Cells[offset:] {
			if hasBorder(cell) {
				r = append(r, getData(cell))
				fmt.Printf("[%d]%q ", nc, r[nc])
			}
		}
		mtrx = append(mtrx, r)
		fmt.Println()
	}

	if len(mtrx) == 0 {
		return nil, errors.New("No data")
	}

	tbl := NewTable(len(mtrx), len(mtrx[0]))
	tbl.data = mtrx

	return tbl, err
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
