package xlsx

import (
	"bytes"
	"errors"
	"fmt"
	"io"
	"io/ioutil"
	"reflect"
	"strings"
	"time"

	"github.com/araddon/dateparse"

	"github.com/bingoohuang/gor"
	"github.com/sirupsen/logrus"
	"github.com/unidoc/unioffice/spreadsheet"
)

// T just for tag for convenience to declare some tags for the whole structure.
type T interface{ t() }

// Xlsx is the structure for xlsx processing.
type Xlsx struct {
	workbook     *spreadsheet.Workbook
	currentSheet spreadsheet.Sheet
	option       *Option
	rowsWritten  int
}

func (x *Xlsx) hasInput() bool { return x.option.TemplateFile != "" || x.option.InputFile != "" }

// New creates a new instance of Xlsx.
func New(optionFns ...OptionFn) (x *Xlsx, err error) {
	x = &Xlsx{option: createOption(optionFns)}

	if err := x.readFile(); err != nil {
		return nil, err
	}

	if err := x.readerExcel(); err != nil {
		return nil, err
	}

	if x.workbook == nil {
		x.workbook = spreadsheet.New()
	}

	return x, nil
}

func (x *Xlsx) readFile() (err error) {
	if t := x.option.TemplateFile; t != "" {
		if x.workbook, err = spreadsheet.Open(t); err != nil {
			logrus.Warnf("failed to open template file %s: %v", t, err)
			return err
		}
	}

	if t := x.option.InputFile; t != "" {
		if x.workbook, err = spreadsheet.Open(t); err != nil {
			logrus.Warnf("failed to open input file %s: %v", t, err)
			return err
		}
	}

	if t := x.option.httpUpload; t != nil {
		if x.workbook, err = t.parseUploadFile(); err != nil {
			logrus.Warnf("failed to parseUploadFile for the file key %s: %v", t.filenameKey, err)
			return err
		}
	}

	return nil
}

func (x *Xlsx) readerExcel() error {
	t := x.option.Reader
	if t == nil {
		return nil
	}

	readerBytes, err := ioutil.ReadAll(t)
	if err != nil {
		return err
	}

	r := bytes.NewReader(readerBytes)
	if x.workbook, err = spreadsheet.Read(r, int64(len(readerBytes))); err != nil {
		logrus.Warnf("failed to read input file from the reader: %v", err)
		return err
	}

	return nil
}

func createOption(optionFns []OptionFn) *Option {
	option := &Option{}

	for _, fn := range optionFns {
		fn(option)
	}

	return option
}

// Option defines the option for the xlsx processing.
type Option struct {
	TemplateFile string
	InputFile    string
	httpUpload   *upload
	Validations  map[string][]string
	Reader       io.Reader
	Placeholder  bool
}

// OptionFn defines the func to change the option.
type OptionFn func(*Option)

// WithTemplate defines the template excel file for writing template.
func WithTemplate(f string) OptionFn { return func(o *Option) { o.TemplateFile = f } }

// WithTemplatePlaceholder defines the template excel file for writing template in placeholder mode.
func WithTemplatePlaceholder(f string) OptionFn {
	return func(o *Option) { o.TemplateFile = f; o.Placeholder = true }
}

// WithInputFile defines the input excel file for reading.
func WithInputFile(f string) OptionFn { return func(o *Option) { o.InputFile = f } }

// WithValidations defines the validations for the cells.
func WithValidations(v map[string][]string) OptionFn { return func(o *Option) { o.Validations = v } }

// WithReader defines the io reader for the writing template or reading excel.
func WithReader(v io.Reader) OptionFn { return func(o *Option) { o.Reader = v } }

// Close does some cleanup like remove temporary files.
func (x *Xlsx) Close() error {
	return x.workbook.Close()
}

// Write Writes beans to the underlying xlsx.
func (x *Xlsx) Write(beans interface{}) error {
	beanReflectValue := reflect.ValueOf(beans)
	beanType := beanReflectValue.Type()
	isSlice := beanReflectValue.Kind() == reflect.Slice

	if isSlice {
		if beanReflectValue.Len() == 0 {
			return nil
		}

		beanType = beanReflectValue.Type().Elem()
	}

	if beanType.Kind() == reflect.Ptr {
		beanType = beanType.Elem()
	}

	ttag := findTTag(beanType)
	x.currentSheet = x.createSheet(ttag, false)

	fields := collectExportableFields(beanType)

	if x.option.Placeholder {
		x.processPlaceholders(beanReflectValue, isSlice, fields)

		return nil
	}

	titles, customizedTitle := collectTitles(fields)
	location := x.locateTitleRow(fields, titles, false)
	customizedTitle = customizedTitle && !location.isValid()

	if writeTitle := customizedTitle || ttag.Get("title") != ""; writeTitle {
		x.writeTitles(fields, titles)
	}

	if location.isValid() {
		x.rowsWritten = 0

		if isSlice {
			for i := 0; i < beanReflectValue.Len(); i++ {
				x.writeTemplateRow(location, beanReflectValue.Index(i))
			}
		} else {
			x.writeTemplateRow(location, beanReflectValue)
		}

		x.removeTempleRows(location)

		return x.createTemplateDataValidations(location, x.currentSheet)
	}

	if isSlice {
		for i := 0; i < beanReflectValue.Len(); i++ {
			x.writeRow(fields, beanReflectValue.Index(i))
		}
	} else {
		x.writeRow(fields, beanReflectValue)
	}

	return x.createDataValidations(fields, x.currentSheet)
}

func (x *Xlsx) processPlaceholders(beanReflectValue reflect.Value, isSlice bool, fields []reflect.StructField) {
	placeholders := collectPlaceholders(x.currentSheet)
	v := beanReflectValue

	if isSlice {
		v = beanReflectValue.Index(0)
	}

	x.writePlaceholderTemplate(fields, placeholders, v)
}

func (x *Xlsx) writePlaceholderTemplate(fields []reflect.StructField,
	plMap map[string]PlaceholderValue, v reflect.Value) {
	vars := make(map[string]string)
	placeholderCells := make(map[string]string)

	for _, f := range fields {
		name := f.Tag.Get("placeholder")
		if name == "" {
			name = f.Name
		}

		vars[name] = getFieldValue(f, v)

		if v := f.Tag.Get("placeholderCell"); v != "" {
			placeholderCells[v] = vars[name]
		}
	}

	for k, v := range plMap {
		x.currentSheet.Cell(k).SetString(v.Interpolate(vars))
	}

	for k, v := range placeholderCells {
		x.currentSheet.Cell(k).SetString(v)
	}
}

func collectPlaceholders(sheet spreadsheet.Sheet) map[string]PlaceholderValue {
	placeholders := make(map[string]PlaceholderValue)

	for _, row := range sheet.Rows() {
		for _, cell := range row.Cells() {
			if pl := ParsePlaceholder(cell.GetString()); pl.HasPlaceholders() {
				placeholders[cell.Reference()] = pl
			}
		}
	}

	return placeholders
}

// PlaceholderValue represents a placeholder value.
type PlaceholderValue struct {
	PlaceholderVars   []string
	PlaceholderQuotes []string
	Content           string
}

// HasPlaceholders tells that the PlaceholderValue has any placeholders.
func (p *PlaceholderValue) HasPlaceholders() bool { return len(p.PlaceholderVars) > 0 }

// Interpolate interpolates placeholders with vars.
func (p *PlaceholderValue) Interpolate(vars map[string]string) string {
	content := p.Content

	for i := 0; i < len(p.PlaceholderVars); i++ {
		v := vars[p.PlaceholderVars[i]]
		content = strings.ReplaceAll(content, p.PlaceholderQuotes[i], v)
	}

	return content
}

// ParsePlaceholder parses placeholders in the content.
func ParsePlaceholder(content string) PlaceholderValue {
	placeholders := make([]string, 0)
	placeholderQuotes := make([]string, 0)

	pos := 0

	for {
		lp := strings.Index(content[pos:], "{{")
		if lp < 0 {
			break
		}

		rp := strings.Index(content[pos+lp:], "}}")
		if rp < 0 {
			break
		}

		pl := content[pos+lp : pos+lp+rp+2]
		placeholderQuotes = append(placeholderQuotes, pl)
		placeholders = append(placeholders, strings.TrimSpace(pl[2:len(pl)-2]))

		pos += lp + rp
	}

	return PlaceholderValue{
		PlaceholderVars:   placeholders,
		PlaceholderQuotes: placeholderQuotes,
		Content:           content,
	}
}

func (x *Xlsx) createDataValidations(fields []reflect.StructField, sheet spreadsheet.Sheet) error {
	for i, field := range fields {
		cellColumn, _ := sheet.Rows()[0].Cells()[i].Column()

		dv := field.Tag.Get("dataValidation")
		// nolint gomnd
		if err := x.createColumnDataValidation(2, sheet, dv, cellColumn); err != nil {
			return err
		}
	}

	return nil
}

func (x *Xlsx) createTemplateDataValidations(l templateLocation, sheet spreadsheet.Sheet) error {
	for _, tc := range l.templateCells {
		cellColumn := tc.cellColumn

		dv := tc.structField.Tag.Get("dataValidation")
		// nolint gomnd
		if err := x.createColumnDataValidation(l.titledRowNumber+1, sheet, dv, cellColumn); err != nil {
			return err
		}
	}

	return nil
}

func (x *Xlsx) createColumnDataValidation(startRowNum int, sheet spreadsheet.Sheet, tag, cellColumn string) error {
	if tag == "" {
		return nil
	}

	dvCombo := sheet.AddDataValidation()
	rangeRef := fmt.Sprintf("%s%d:%s%d", cellColumn, startRowNum, cellColumn, startRowNum+x.rowsWritten)

	dvCombo.SetRange(rangeRef)

	dvList := dvCombo.SetList()

	if strings.Contains(tag, "!") {
		dvs := strings.Split(tag, "!")
		dvSheetName, validateRange := dvs[0], dvs[1]
		vsheet := x.findSheet(dvSheetName)

		if !vsheet.IsValid() {
			return fmt.Errorf("unable to find sheet with name %s", dvSheetName)
		}

		dvList.SetRange(vsheet.RangeReference(validateRange))
	} else {
		if vm, ok := x.option.Validations[tag]; ok {
			dvList.SetValues(vm)
		} else {
			dvList.SetValues(strings.Split(tag, ","))
		}
	}

	return nil
}

func (x *Xlsx) Read(slicePtr interface{}) error {
	v := reflect.ValueOf(slicePtr)
	if v.Kind() != reflect.Ptr || v.Elem().Kind() != reflect.Slice {
		return errors.New("the input argument should be a pointer of slice")
	}

	beanType := v.Elem().Type().Elem()

	ttag := findTTag(beanType)
	x.currentSheet = x.createSheet(ttag, true)

	fields := collectExportableFields(beanType)
	titles, _ := collectTitles(fields)
	location := x.locateTitleRow(fields, titles, true)

	if location.isValid() {
		slice, err := x.readRows(beanType, location)
		if err != nil {
			return err
		}

		v.Elem().Set(slice)
	}

	return nil
}

func (x *Xlsx) readRows(beanType reflect.Type, l templateLocation) (reflect.Value, error) {
	slice := reflect.MakeSlice(reflect.SliceOf(beanType), 0, len(l.templateRows))

	for _, row := range l.templateRows {
		rowBean, err := x.createRowBean(beanType, l, row)
		if err != nil {
			return reflect.Value{}, err
		}

		slice = reflect.Append(slice, rowBean)
	}

	return slice, nil
}

func (x *Xlsx) createRowBean(beanType reflect.Type, l templateLocation, row spreadsheet.Row) (reflect.Value, error) {
	rowBean := reflect.New(beanType).Elem()

	for _, cell := range l.templateCells {
		c := row.Cell(cell.cellColumn)
		if c.IsEmpty() {
			continue
		}

		sf := cell.structField
		f := rowBean.FieldByIndex(sf.Index)
		s := c.GetString()

		if sf.Type == timeType {
			t, err := parseTime(sf, s)
			if err != nil {
				return reflect.Value{}, err
			}

			f.Set(reflect.ValueOf(t))

			continue
		}

		v, err := gor.CastAny(s, sf.Type)

		if err != nil && sf.Tag.Get("omiterr") != "true" {
			return reflect.Value{}, err
		}

		f.Set(v)
	}

	return rowBean, nil
}

func parseTime(sf reflect.StructField, s string) (time.Time, error) {
	if f := sf.Tag.Get("format"); f != "" {
		return time.ParseInLocation(ParseJavaTimeFormat(f), s, time.Local)
	}

	return dateparse.ParseLocal(s)
}

func (x *Xlsx) createSheet(ttag reflect.StructTag, readonly bool) spreadsheet.Sheet {
	sheetName := ttag.Get("sheet")
	wbSheet := spreadsheet.Sheet{}

	if x.hasInput() {
		if sh := x.findSheet(sheetName); sh.IsValid() {
			return sh
		}

		if len(x.workbook.Sheets()) > 0 {
			wbSheet = x.workbook.Sheets()[0]
		}
	}

	if readonly {
		return wbSheet
	}

	if !wbSheet.IsValid() {
		wbSheet = x.workbook.AddSheet()
	}

	if sheetName != "" && !strings.Contains(wbSheet.Name(), sheetName) {
		wbSheet.SetName(sheetName)
	}

	return wbSheet
}

func collectTitles(fields []reflect.StructField) ([]string, bool) {
	titles := make([]string, 0)
	customizedTitle := false

	for _, f := range fields {
		if t := f.Tag.Get("title"); t != "" {
			customizedTitle = true

			titles = append(titles, t)
		} else {
			titles = append(titles, f.Name)
		}
	}

	return titles, customizedTitle
}

func collectExportableFields(t reflect.Type) []reflect.StructField {
	fields := make([]reflect.StructField, 0, t.NumField())

	for i := 0; i < t.NumField(); i++ {
		f := t.Field(i)

		if f.PkgPath != "" || f.Type == tType {
			continue
		}

		fields = append(fields, f)
	}

	return fields
}

func findTTag(t reflect.Type) reflect.StructTag {
	for i := 0; i < t.NumField(); i++ {
		if f := t.Field(i); f.Type == tType {
			return f.Tag
		}
	}

	return ""
}

// SaveToFile writes the workbook out to a file.
func (x *Xlsx) SaveToFile(file string) error { return x.workbook.SaveToFile(file) }

// Save writes the workbook out to a writer in the zipped xlsx format.
func (x *Xlsx) Save(w io.Writer) error { return x.workbook.Save(w) }

func (x *Xlsx) writeRow(fields []reflect.StructField, value reflect.Value) {
	row := x.currentSheet.AddRow()
	x.rowsWritten++

	for _, field := range fields {
		setCellValue(row.AddCell(), field, value)
	}
}

func getFieldValue(field reflect.StructField, value reflect.Value) string {
	v := value.FieldByIndex(field.Index).Interface()

	if fv, ok := ConvertNumberToFloat64(v); ok {
		return fmt.Sprintf("%v", fv)
	}

	switch fv := v.(type) {
	case time.Time:
		if format := field.Tag.Get("format"); format != "" {
			return fv.Format(ParseJavaTimeFormat(format))
		}

		return fv.Format("2006-01-02 15:04:05")
	case string:
		return fv
	case bool:
		return fmt.Sprintf("%v", fv)
	case nil:
		return ""
	default:
		return ""
	}
}

func setCellValue(cell spreadsheet.Cell, field reflect.StructField, value reflect.Value) {
	v := value.FieldByIndex(field.Index).Interface()

	if fv, ok := ConvertNumberToFloat64(v); ok {
		cell.SetNumber(fv)
		return
	}

	switch fv := v.(type) {
	case time.Time:
		if format := field.Tag.Get("format"); format != "" {
			cell.SetString(fv.Format(ParseJavaTimeFormat(format)))
		} else {
			cell.SetTime(fv)
		}
	case string:
		cell.SetString(fv)
	case bool:
		cell.SetBool(fv)
	case nil:
		cell.SetString("")
	}
}

func (x *Xlsx) writeTitles(fields []reflect.StructField, titles []string) {
	row := x.currentSheet.AddRow()

	for i := range fields {
		row.AddCell().SetString(titles[i])
	}
}

type templateLocation struct {
	titledRowNumber int
	rowsEndIndex    int
	templateCells   []templateCell
	templateRows    []spreadsheet.Row
}

type templateCell struct {
	cellColumn  string
	structField reflect.StructField
}

func (t *templateLocation) isValid() bool {
	return len(t.templateCells) > 0
}

func (x *Xlsx) locateTitleRow(fields []reflect.StructField, titles []string, forRead bool) templateLocation {
	if !x.hasInput() {
		return templateLocation{}
	}

	rows := x.currentSheet.Rows()
	titledRowNumber, templateCells := x.findTemplateTitledRow(fields, titles, rows)
	templateRows := x.findTemplateRows(titledRowNumber, templateCells, rows, forRead)

	return templateLocation{
		titledRowNumber: titledRowNumber,
		templateRows:    templateRows,
		templateCells:   templateCells,
		rowsEndIndex:    len(rows),
	}
}

func (x *Xlsx) findTemplateTitledRow(fields []reflect.StructField,
	titles []string, rows []spreadsheet.Row) (int, []templateCell) {
	titledRowNumber := -1
	templateCells := make([]templateCell, 0, len(fields))

	for _, row := range rows {
		for _, cell := range row.Cells() {
			cellString := cell.GetString()

			for i, title := range titles {
				if !strings.Contains(cellString, title) {
					continue
				}

				col, err := cell.Column()
				if err != nil {
					logrus.Warnf("failed to get column error: %v", err)
					continue
				}

				templateCells = append(templateCells, templateCell{
					cellColumn:  col,
					structField: fields[i],
				})

				break
			}
		}

		if len(templateCells) > 0 {
			titledRowNumber = int(row.RowNumber())
			break
		}
	}

	return titledRowNumber, templateCells
}

func (x *Xlsx) findTemplateRows(titledRowNumber int,
	templateCells []templateCell, rows []spreadsheet.Row, forRead bool) []spreadsheet.Row {
	templateRows := make([]spreadsheet.Row, 0)

	if titledRowNumber < 0 {
		return templateRows
	}

	col := templateCells[0].cellColumn

	for i := 0; i < len(rows); i++ {
		if rows[i].RowNumber() <= uint32(titledRowNumber) {
			continue
		}

		if forRead || strings.Contains(rows[i].Cell(col).GetString(), "template") {
			templateRows = append(templateRows, rows[i])
		} else if len(templateRows) == 0 {
			return append(templateRows, rows[i])
		}
	}

	return templateRows
}

func (x *Xlsx) writeTemplateRow(l templateLocation, v reflect.Value) {
	// 2 是为了计算row num(1-N), 从标题行(T)的下一行（T+1)开始写
	num := uint32(l.titledRowNumber + 1 + x.rowsWritten)
	x.rowsWritten++
	row := x.currentSheet.Row(num)

	for _, tc := range l.templateCells {
		setCellValue(row.Cell(tc.cellColumn), tc.structField, v)
	}

	x.copyRowStyle(l, row)
}

func (x *Xlsx) copyRowStyle(l templateLocation, row spreadsheet.Row) {
	if len(l.templateRows) == 0 || x.rowsWritten < len(l.templateRows) {
		return
	}

	templateRow := l.templateRows[(x.rowsWritten-1)%len(l.templateRows)]

	for _, tc := range l.templateCells { // copying cell style
		cell := templateRow.Cell(tc.cellColumn)
		if cx := cell.X(); cx.SAttr != nil {
			if style := x.workbook.StyleSheet.GetCellStyle(*cx.SAttr); !style.IsEmpty() {
				row.Cell(tc.cellColumn).SetStyle(style)
			}
		}
	}
}

func (x *Xlsx) removeTempleRows(l templateLocation) {
	if len(l.templateRows) == 0 {
		return // no template rows in the template file.
	}

	sheetData := x.currentSheet.X().CT_Worksheet.SheetData
	rows := sheetData.Row

	if endIndex := l.titledRowNumber + x.rowsWritten; endIndex < len(rows) {
		sheetData.Row = rows[:endIndex]
	}
}

func (x *Xlsx) findSheet(sheetName string) spreadsheet.Sheet {
	for _, sheet := range x.workbook.Sheets() {
		if strings.Contains(sheet.Name(), sheetName) {
			return sheet
		}
	}

	return spreadsheet.Sheet{}
}

// nolint gochecknoglobals
var (
	tType    = reflect.TypeOf((*T)(nil)).Elem()
	timeType = reflect.TypeOf((*time.Time)(nil)).Elem()
)

// ParseJavaTimeFormat converts the time format in java to golang.
func ParseJavaTimeFormat(layout string) string {
	lo := layout
	lo = strings.Replace(lo, "yyyy", "2006", -1)
	lo = strings.Replace(lo, "yy", "06", -1)
	lo = strings.Replace(lo, "MM", "01", -1)
	lo = strings.Replace(lo, "dd", "02", -1)
	lo = strings.Replace(lo, "HH", "15", -1)
	lo = strings.Replace(lo, "mm", "04", -1)
	lo = strings.Replace(lo, "ss", "05", -1)
	lo = strings.Replace(lo, "SSS", "000", -1)

	return lo
}

// ConvertNumberToFloat64 converts a number value to float64.
// If the value is not a number, it returns 0, false.
func ConvertNumberToFloat64(v interface{}) (float64, bool) {
	switch fv := v.(type) {
	case int:
		return float64(fv), true
	case int8:
		return float64(fv), true
	case int16:
		return float64(fv), true
	case int32:
		return float64(fv), true
	case int64:
		return float64(fv), true
	case uint:
		return float64(fv), true
	case uint8:
		return float64(fv), true
	case uint16:
		return float64(fv), true
	case uint32:
		return float64(fv), true
	case uint64:
		return float64(fv), true
	case float32:
		return float64(fv), true
	case float64:
		return fv, true
	}

	return 0, false
}
