package xlsx

import (
	"bytes"
	"io"
	"io/ioutil"

	"github.com/sirupsen/logrus"
	"github.com/unidoc/unioffice/spreadsheet"
)

func createOption(optionFns []OptionFn) *Option {
	option := &Option{}

	for _, fn := range optionFns {
		fn(option)
	}

	return option
}

// Option defines the option for the xlsx processing.
type Option struct {
	TemplateWorkbook *spreadsheet.Workbook
	Workbook         *spreadsheet.Workbook
	httpUpload       *upload
	Validations      map[string][]string
	Placeholder      bool
}

// OptionFn defines the func to change the option.
type OptionFn func(*Option)

// WithTemplateFile defines the template excel file for writing template.
func WithTemplateFile(f string) OptionFn {
	wb, err := spreadsheet.Open(f)
	if err != nil {
		logrus.Warnf("failed to open template file %s: %v", f, err)
		return func(*Option) {}
	}

	return func(o *Option) { o.TemplateWorkbook = wb }
}

// WithTemplateBytes defines the template excel file for writing template.
func WithTemplateBytes(f []byte) OptionFn {
	wb, err := spreadsheet.Read(bytes.NewReader(f), int64(len(f)))
	if err != nil {
		logrus.Warnf("failed to open file %s: %v", f, err)
		return func(*Option) {}
	}

	return func(o *Option) { o.TemplateWorkbook = wb }
}

// WithTemplateReader defines the template excel file for writing template.
func WithTemplateReader(f io.Reader) OptionFn {
	readerBytes, err := ioutil.ReadAll(f)
	if err != nil {
		logrus.Warnf("failed to ReadAll: %v", err)
		return func(*Option) {}
	}

	return WithTemplateBytes(readerBytes)
}

// AsPlaceholder defines the template excel file for writing template in placeholder mode.
func AsPlaceholder() OptionFn {
	return func(o *Option) { o.Placeholder = true }
}

// WithFile defines the input excel file for reading.
func WithFile(f string) OptionFn {
	wb, err := spreadsheet.Open(f)
	if err != nil {
		logrus.Warnf("failed to open file %s: %v", f, err)
		return func(*Option) {}
	}

	return func(o *Option) { o.Workbook = wb }
}

// WithBytes defines the input excel file bytes for reading.
func WithBytes(f []byte) OptionFn {
	wb, err := spreadsheet.Read(bytes.NewReader(f), int64(len(f)))
	if err != nil {
		logrus.Warnf("failed to open file %s: %v", f, err)
		return func(*Option) {}
	}

	return func(o *Option) { o.Workbook = wb }
}

// WithReader defines the input excel file reader for reading.
func WithReader(f io.Reader) OptionFn {
	readerBytes, err := ioutil.ReadAll(f)
	if err != nil {
		logrus.Warnf("failed to ReadAll: %v", err)
		return func(*Option) {}
	}

	return WithBytes(readerBytes)
}

// WithValidations defines the validations for the cells.
func WithValidations(v map[string][]string) OptionFn {
	return func(o *Option) { o.Validations = v }
}
