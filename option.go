package xlsx

import (
	"bytes"
	"fmt"
	"io"
	"io/ioutil"

	"github.com/sirupsen/logrus"
	"github.com/unidoc/unioffice/spreadsheet"
)

func createOption(optionFns []OptionFn) *Option {
	option := &Option{}

	for _, fn := range optionFns {
		if fn != nil {
			fn(option)
		}
	}

	return option
}

// Option defines the option for the xlsx processing.
type Option struct {
	TemplateWorkbook, Workbook *spreadsheet.Workbook

	Validations map[string][]string
}

// OptionFn defines the func to change the option.
type OptionFn func(*Option)

// WithTemplate defines the template excel file for writing template.
// The template can be type of any of followings:
// 1. a string for direct template excel file name
// 2. a []byte for the content of template excel which loaded in advance, like use packr2 to read.
// 3. a io.Reader.
func WithTemplate(template interface{}) OptionFn {
	wb, err := parseExcel(template)
	if err != nil {
		logrus.Warnf("failed to open template excel %v", err)
		return nil
	}

	return func(o *Option) { o.TemplateWorkbook = wb }
}

// WithExcel defines the input excel file for reading.
// The excel can be type of any of followings:
// 1. a string for direct excel file name
// 2. a []byte for the content of excel which loaded in advance, like use packr2 to read.
// 3. a io.Reader.
func WithExcel(excel interface{}) OptionFn {
	wb, err := parseExcel(excel)
	if err != nil {
		logrus.Warnf("failed to open excel %v", err)
		return nil
	}

	return func(o *Option) { o.Workbook = wb }
}

// UnknownExcelError defines the the unknown excel file format error.
var UnknownExcelError = fmt.Errorf("unknown excel file format")

func parseExcel(f interface{}) (wb *spreadsheet.Workbook, err error) {
	var bs []byte

	switch ft := f.(type) {
	case string:
		return spreadsheet.Open(ft)
	case []byte:
		return spreadsheet.Read(bytes.NewReader(ft), int64(len(ft)))
	case io.Reader:
		if bs, err = ioutil.ReadAll(ft); err != nil {
			return nil, err
		}

		return spreadsheet.Read(bytes.NewReader(bs), int64(len(bs)))
	default:
		return nil, UnknownExcelError
	}
}

// WithValidations defines the validations for the cells.
func WithValidations(v map[string][]string) OptionFn {
	return func(o *Option) { o.Validations = v }
}
