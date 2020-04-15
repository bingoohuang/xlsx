package xlsx

import (
	"errors"
	"fmt"
	"io"
	"reflect"
	"strings"
	"time"

	"github.com/araddon/dateparse"

	"github.com/bingoohuang/gor"
	"github.com/sirupsen/logrus"
	"github.com/unidoc/unioffice/spreadsheet"
)

// T just for tag for convenience to declare some tags for the whole structure.
type T struct{}

// Xlsx is the structure for xlsx processing.
type Xlsx struct {
	tmplWorkbook, workbook  *spreadsheet.Workbook
	tmplSheet, currentSheet spreadsheet.Sheet
	option                  *Option
	rowsWritten             int
}

func (x *Xlsx) hasInput() bool { return x.option.TemplateWorkbook != nil || x.option.Workbook != nil }

// New creates a new instance of Xlsx.
func New(optionFns ...OptionFn) (x *Xlsx, err error) {
	x = &Xlsx{option: createOption(optionFns)}

	x.tmplWorkbook = x.option.TemplateWorkbook
	x.workbook = x.option.Workbook

	if x.workbook == nil && x.tmplWorkbook != nil {
		x.workbook = x.tmplWorkbook
		x.tmplWorkbook = nil
	}

	if x.workbook == nil {
		x.workbook = spreadsheet.New()
	}

	return x, nil
}

// Close does some cleanup like remove temporary files.
func (x *Xlsx) Close() error {
	if x.tmplWorkbook != nil {
		_ = x.tmplWorkbook.Close()
	}

	return x.workbook.Close()
}

type run struct {
	isSlice   bool
	isPtr     bool
	beanType  reflect.Type
	rawValue  reflect.Value
	beanValue reflect.Value
	ttags     []reflect.StructTag
	fields    []reflect.StructField
}

func makeRun(beans interface{}) *run {
	r := &run{}
	r.rawValue = reflect.ValueOf(beans)
	r.beanValue = r.rawValue
	r.beanType = r.beanValue.Type()
	r.isPtr = r.beanType.Kind() == reflect.Ptr

	if r.isPtr {
		r.beanType = r.beanType.Elem()
		r.beanValue = r.beanValue.Elem()
	}

	r.isSlice = r.beanValue.Kind() == reflect.Slice

	if r.isSlice {
		r.beanType = r.beanType.Elem()
	}

	r.collectTtags()
	r.collectExportableFields()

	return r
}

func (r *run) isEmptySlice() bool {
	return r.isSlice && r.beanValue.Len() == 0
}

func (r *run) collectTtags() {
	r.ttags = make([]reflect.StructTag, 0)

	for i := 0; i < r.beanType.NumField(); i++ {
		if f := r.beanType.Field(i); f.Type == tType {
			r.ttags = append(r.ttags, f.Tag)
		}
	}
}

func (r *run) FindTtag(tagName string) string {
	for _, t := range r.ttags {
		if v := t.Get(tagName); v != "" {
			return v
		}
	}

	return ""
}

func (r *run) collectExportableFields() {
	t := r.beanType
	r.fields = make([]reflect.StructField, 0, t.NumField())

	for i := 0; i < t.NumField(); i++ {
		f := t.Field(i)

		if f.PkgPath != "" || f.Type == tType {
			continue
		}

		r.fields = append(r.fields, f)
	}
}

func (r *run) getSingleBean() reflect.Value {
	if r.isSlice {
		return r.beanValue.Index(0)
	}

	return r.beanValue
}

func (r *run) forRead() bool {
	return r.isPtr && (r.isSlice || r.beanType.Kind() == reflect.Struct)
}

func (r *run) AsPlaceholderTtag() bool {
	v := r.FindTtag("asPlaceholder")
	return v == "true" || v == "on" || v == "yes" || v == "1"
}

// Write Writes beans to the underlying xlsx.
func (x *Xlsx) Write(beans interface{}) error {
	r := makeRun(beans)
	if r.isEmptySlice() {
		return nil
	}

	x.currentSheet = x.createSheet(x.workbook, r, false)

	if r.AsPlaceholderTtag() {
		x.writePlaceholderTemplate(r.fields, collectPlaceholders(x.currentSheet), r.getSingleBean())

		return nil
	}

	titles, customizedTitle := collectTitles(r.fields)
	location := x.locateTitleRow(r.fields, titles)
	customizedTitle = customizedTitle && !location.isValid()

	if writeTitles := customizedTitle || r.FindTtag("title") != ""; writeTitles {
		x.writeTitles(r.fields, titles)
	}

	if location.isValid() {
		x.rowsWritten = 0

		if r.isSlice {
			for i := 0; i < r.beanValue.Len(); i++ {
				x.writeTemplateRow(location, r.beanValue.Index(i))
			}
		} else {
			x.writeTemplateRow(location, r.beanValue)
		}

		x.removeTempleRows(location)

		return x.createTemplateDataValidations(location, x.currentSheet)
	}

	if r.isSlice {
		for i := 0; i < r.beanValue.Len(); i++ {
			x.writeRow(r.fields, r.beanValue.Index(i))
		}
	} else {
		x.writeRow(r.fields, r.beanValue)
	}

	return x.createDataValidations(r.fields, x.currentSheet)
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

// nolint gomnd
func (x *Xlsx) createDataValidations(fields []reflect.StructField, sheet spreadsheet.Sheet) error {
	row0Cells := sheet.Rows()[0].Cells()

	for i, field := range fields {
		cellColumn, _ := row0Cells[i].Column()

		dv := field.Tag.Get("dataValidation")
		if err := x.createColumnDataValidation(2, sheet, dv, cellColumn); err != nil {
			return err
		}
	}

	return nil
}

func (x *Xlsx) createTemplateDataValidations(l templateLocation, sheet spreadsheet.Sheet) error {
	for _, tc := range l.templateCells {
		dv := tc.structField.Tag.Get("dataValidation")
		// nolint gomnd
		if err := x.createColumnDataValidation(l.titledRowNumber+1, sheet, dv, tc.cellColumn); err != nil {
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
		vsheet := x.findSheet(x.workbook, dvSheetName)

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
	r := makeRun(slicePtr)

	if !r.forRead() {
		return errors.New("the input argument should be a pointer of slice")
	}

	x.tmplSheet = x.createSheet(x.tmplWorkbook, r, true)
	x.currentSheet = x.createSheet(x.workbook, r, true)

	if r.AsPlaceholderTtag() {
		err := x.writePlaceholderToBean(r)
		if err != nil {
			return err
		}

		return nil
	}

	titles, _ := collectTitles(r.fields)
	location := x.locateTitleRow(r.fields, titles)

	if location.isValid() {
		slice, err := x.readRows(r.beanType, location)
		if err != nil {
			return err
		}

		r.rawValue.Elem().Set(slice)
	}

	return nil
}

func (x *Xlsx) writePlaceholderToBean(r *run) error {
	vars := x.readPlaceholderValues()
	vv := r.beanValue

	for _, f := range r.fields {
		if v := f.Tag.Get("placeholderCell"); v != "" {
			vs := x.currentSheet.Cell(v).GetString()
			if err := setFieldValue(vv, f, vs); err != nil {
				return err
			}

			continue
		}

		name := f.Tag.Get("placeholder")
		if name == "" {
			name = f.Name
		}

		if varValue, ok := vars[name]; ok {
			if err := setFieldValue(vv, f, varValue); err != nil {
				return err
			}
		}
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

		if err := setFieldValue(rowBean, cell.structField, c.GetString()); err != nil {
			return reflect.Value{}, err
		}
	}

	return rowBean, nil
}

func setFieldValue(rowBean reflect.Value, sf reflect.StructField, s string) error {
	f := rowBean.FieldByIndex(sf.Index)

	if sf.Type == timeType {
		t, err := parseTime(sf.Tag, s)
		if err != nil {
			return err
		}

		f.Set(reflect.ValueOf(t))

		return nil
	}

	v, err := gor.CastAny(s, sf.Type)

	if err != nil && sf.Tag.Get("omiterr") != "true" {
		return err
	}

	f.Set(v)

	return nil
}

func (x *Xlsx) createSheet(wb *spreadsheet.Workbook, r *run, readonly bool) spreadsheet.Sheet {
	wbSheet := spreadsheet.Sheet{}

	if wb == nil {
		return wbSheet
	}

	sheetName := r.FindTtag("sheet")

	if x.hasInput() {
		if sh := x.findSheet(wb, sheetName); sh.IsValid() {
			return sh
		}

		if len(wb.Sheets()) > 0 {
			wbSheet = wb.Sheets()[0]
		}
	}

	if readonly {
		return wbSheet
	}

	if !wbSheet.IsValid() {
		wbSheet = wb.AddSheet()
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
		return formatTime(field.Tag, fv)
	case nil:
		return ""
	default:
		return fmt.Sprintf("%v", fv)
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
		cell.SetString(formatTime(field.Tag, fv))
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

func (x *Xlsx) locateTitleRow(fields []reflect.StructField, titles []string) templateLocation {
	if !x.hasInput() {
		return templateLocation{}
	}

	rows := x.currentSheet.Rows()
	titledRowNumber, templateCells := x.findTemplateTitledRow(fields, titles, rows)
	templateRows := x.findTemplateRows(titledRowNumber, rows)

	return templateLocation{
		titledRowNumber: titledRowNumber,
		templateRows:    templateRows,
		templateCells:   templateCells,
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

func (x *Xlsx) findTemplateRows(titledRowNumber int, rows []spreadsheet.Row) []spreadsheet.Row {
	templateRows := make([]spreadsheet.Row, 0)

	if titledRowNumber < 0 {
		return templateRows
	}

	for _, row := range rows {
		if row.RowNumber() > uint32(titledRowNumber) {
			templateRows = append(templateRows, row)
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
		if cx := templateRow.Cell(tc.cellColumn).X(); cx.SAttr != nil {
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

func (x *Xlsx) findSheet(wb *spreadsheet.Workbook, sheetName string) spreadsheet.Sheet {
	for _, sheet := range wb.Sheets() {
		if strings.Contains(sheet.Name(), sheetName) {
			return sheet
		}
	}

	return spreadsheet.Sheet{}
}

func (x *Xlsx) readPlaceholderValues() map[string]string {
	plMap := collectPlaceholders(x.tmplSheet)
	plVars := make(map[string]string)

	for k, v := range plMap {
		cellValue := x.currentSheet.Cell(k).GetString()

		if vars, ok := v.ParseVars(cellValue); ok {
			for vk, vv := range vars {
				plVars[vk] = vv
			}
		} else {
			logrus.Warnf("failed to parse vars from cell value %s by Part %+v", cellValue, v.Content)
		}
	}

	return plVars
}

// nolint gochecknoglobals
var (
	tType    = reflect.TypeOf((*T)(nil)).Elem()
	timeType = reflect.TypeOf((*time.Time)(nil)).Elem()
)

// ParseJavaTimeFormat converts the time format in java to golang.
func ParseJavaTimeFormat(layout string) string {
	l := layout
	l = strings.Replace(l, "yyyy", "2006", -1)
	l = strings.Replace(l, "yy", "06", -1)
	l = strings.Replace(l, "MM", "01", -1)
	l = strings.Replace(l, "dd", "02", -1)
	l = strings.Replace(l, "HH", "15", -1)
	l = strings.Replace(l, "mm", "04", -1)
	l = strings.Replace(l, "ss", "05", -1)

	return strings.Replace(l, "SSS", "000", -1)
}

// ConvertNumberToFloat64 converts a number value to float64.
// If the value is not a number, it returns 0, false.
func ConvertNumberToFloat64(v interface{}) (float64, bool) {
	switch v.(type) {
	case int, int8, int16, int32, int64, uint, uint8, uint16, uint32, uint64, float32, float64:
		return reflect.ValueOf(v).Convert(reflect.TypeOf(float64(0))).Interface().(float64), true
	}

	return 0, false
}

func formatTime(tag reflect.StructTag, t time.Time) string {
	if v := tag.Get("format"); v != "" {
		return t.Format(ParseJavaTimeFormat(v))
	}

	return t.Format("2006-01-02 15:04:05")
}

func parseTime(tag reflect.StructTag, s string) (time.Time, error) {
	if f := tag.Get("format"); f != "" {
		return time.ParseInLocation(ParseJavaTimeFormat(f), s, time.Local)
	}

	return dateparse.ParseLocal(s)
}
