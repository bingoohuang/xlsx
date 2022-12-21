package xlsx

import (
	"flag"
	"fmt"
	"strconv"
	"strings"
	"unsafe"

	"github.com/unidoc/unioffice"
	"github.com/unidoc/unioffice/schema/soo/sml"
	"github.com/unidoc/unioffice/spreadsheet"
	"github.com/unidoc/unioffice/spreadsheet/reference"
)

// RowCells returns a slice of cells.  The cells can be manipulated, but appending
// to the slice will have no effect.
func RowCells(r spreadsheet.Row) []spreadsheet.Cell {
	var ret []spreadsheet.Cell

	lastIndex := -1

	for _, c := range r.X().C {
		if c.RAttr == nil {
			unioffice.Log("RAttr is nil for a cell, skipping.")
			continue
		}

		ref, err := reference.ParseCellReference(*c.RAttr)
		if err != nil {
			unioffice.Log("RAttr is incorrect for a cell: " + *c.RAttr + ", skipping.")
			continue
		}

		currentIndex := int(ref.ColumnIdx)
		// Add lastIndex >= 0 to fix the Row.Cells method when first cell is not available.
		if lastIndex >= 0 && currentIndex-lastIndex > 1 {
			for col := lastIndex + 1; col < currentIndex; col++ {
				ret = append(ret, r.Cell(reference.IndexToColumn(uint32(col))))
			}
		}

		lastIndex = currentIndex

		ret = append(ret, r.Cell(reference.IndexToColumn(ref.ColumnIdx)))
	}

	return ret
}

func CopyRowStyle(from, to spreadsheet.Row) {
	for _, f := range from.Cells() { // copying cell style
		fcol, _ := f.Column()
		t := to.Cell(fcol)

		CopyCellStyle(f, t)
	}
}

func CopyCellStyle(from, to spreadsheet.Cell) {
	if cx := from.X(); cx.SAttr != nil {
		w := *(**spreadsheet.Workbook)(unsafe.Pointer(&from))
		if style := w.StyleSheet.GetCellStyle(*cx.SAttr); !style.IsEmpty() {
			to.SetStyle(style)
		}
	}
}

// GetCellString returns the string in a cell if it's an inline or string table
// string. Otherwise it returns an empty string.
func GetCellString(c spreadsheet.Cell) string {
	x := c.X()

	switch x.TAttr {
	case sml.ST_CellTypeInlineStr:
		if x.Is != nil && x.Is.T != nil {
			return strings.TrimSpace(*x.Is.T)
		}

		if x.V != nil {
			return strings.TrimSpace(*x.V)
		}
	case sml.ST_CellTypeS:
		if x.V == nil {
			return ""
		}

		id, err := strconv.Atoi(*x.V)
		if err != nil {
			return ""
		}

		s, err := GetSharedString(c, id)
		if err != nil {
			return ""
		}

		return s
	}

	if x.V == nil {
		return ""
	}

	return strings.TrimSpace(*x.V)
}

// GetSharedString retrieves a string from the shared strings table by index.
// nolint:goerr113
func GetSharedString(c spreadsheet.Cell, id int) (string, error) {
	if id < 0 {
		return "", fmt.Errorf("invalid string index %d, must be > 0", id)
	}

	w := *(**spreadsheet.Workbook)(unsafe.Pointer(&c))
	x := w.SharedStrings.X()

	if id > len(x.Si) {
		return "", fmt.Errorf("invalid string index %d, table only has %d values", id, len(x.Si))
	}

	si := x.Si[id]

	if si.T != nil {
		return *si.T, nil
	}

	s := ""
	for _, r := range si.R {
		s += r.T
	}

	return strings.TrimSpace(s), nil
}

type flagNoopValue struct{}

func (*flagNoopValue) String() string   { return "noop" }
func (*flagNoopValue) Set(string) error { return nil }

func hackTestV() {
	if flag.Lookup("test.v") == nil {
		flag.CommandLine.Var(&flagNoopValue{}, "test.v", "test.v")
	}
}
