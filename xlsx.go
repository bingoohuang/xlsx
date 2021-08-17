package xlsx

import (
	"errors"
	"fmt"
	"io"
	"log"
	"reflect"
	"sort"
	"strings"
	"time"

	perr "github.com/pkg/errors"
	"github.com/unidoc/unioffice/spreadsheet/reference"

	"github.com/bingoohuang/xlsx/pkg/cast"

	"github.com/araddon/dateparse"

	"github.com/unidoc/unioffice/spreadsheet"
)

// Xlsx is the structure for xlsx processing.
type Xlsx struct {
	tmplWorkbook, workbook  *spreadsheet.Workbook
	tmplSheet, currentSheet spreadsheet.Sheet
	option                  *Option
	rowsWritten             uint32

	tmplSheetReused bool
}

func (x *Xlsx) hasInput() bool {
	return x.option.TemplateWorkbook != nil || x.option.Workbook != nil
}

// New creates a new instance of Xlsx.
func New(optionFns ...OptionFn) (x *Xlsx, err error) {
	hackTestV()

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
	isSlice     bool
	isPtr       bool
	beanType    reflect.Type
	rawValue    reflect.Value
	beanValue   reflect.Value
	ttags       []reflect.StructTag
	fields      []reflect.StructField
	writeOption WriteOption
}

func makeRun(beans interface{}, writeOptionFns []WriteOptionFn) *run {
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

	r.collectTags()
	r.collectExportableFields()

	r.writeOption = WriteOption{}

	for _, fn := range writeOptionFns {
		fn(&r.writeOption)
	}

	return r
}

func (r *run) isEmptySlice() bool {
	return r.isSlice && r.beanValue.Len() == 0
}

func (r *run) collectTags() {
	r.ttags = make([]reflect.StructTag, 0)

	for i := 0; i < r.beanType.NumField(); i++ {
		f := r.beanType.Field(i)
		r.ttags = append(r.ttags, f.Tag)
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

		if f.PkgPath == "" {
			r.fields = append(r.fields, f)
		}
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

func (r *run) asPlaceholder() bool {
	return ParseBool(r.FindTtag("asPlaceholder"), false)
}

func (r *run) ignoreEmptyRows() bool {
	return ParseBool(r.FindTtag("ignoreEmptyRows"), true)
}

// ParseBool parses the v as  a boolean  when it is any of true,  on, yes or 1.
func ParseBool(v string, defaultValue bool) bool {
	if v == "" {
		return defaultValue
	}

	switch strings.ToLower(v) {
	case "true", "t", "on", "yes", "y", "1":
		return true
	default:
		return false
	}
}

type MergeColsMode int

const (
	// DoNotMerge do not mergeTitled.
	DoNotMerge MergeColsMode = iota
	// MergeCols mergeTitled columns separately.
	// like:
	// a, b, 1
	// a, b, 2
	// c, b, 3
	// will merged to :
	// a, b, 1
	// -, -, 2
	// c, -, 3
	MergeCols
	// MergeColsAlign mergeTitled columns align left merging.
	// like:
	// a, b, 1
	// a, b, 2
	// c, b, 3
	// will merged to :
	// a, b, 1
	// -, -, 2
	// c, b, 3
	MergeColsAlign
)

type WriteOption struct {
	SheetName     string
	MergeColsMode MergeColsMode
}

type WriteOptionFn func(*WriteOption)

func WithMergeColsMode(v MergeColsMode) WriteOptionFn {
	return func(o *WriteOption) {
		o.MergeColsMode = v
	}
}

func WithSheetName(v string) WriteOptionFn {
	return func(o *WriteOption) {
		o.SheetName = v
	}
}

// Write Writes beans to the underlying xlsx.
func (x *Xlsx) Write(beans interface{}, writeOptionFns ...WriteOptionFn) error {
	r := makeRun(beans, writeOptionFns)
	if r.isEmptySlice() {
		return nil
	}

	x.tmplSheet, x.currentSheet = x.createWriteSheet(x.workbook, r)

	if r.asPlaceholder() {
		x.writePlaceholder(r.fields, collectPlaceholders(x.currentSheet), r.getSingleBean())

		return nil
	}

	titles, customizedTitles := collectTitles(r.fields)
	loc, err := x.locateTitleRow(titles, customizedTitles)

	if err != nil && !errors.Is(err, ErrNoExcelRead) {
		return err
	}

	location := *loc
	newSheet := x.tmplSheet != x.currentSheet

	if newSheet {
		copyRowsUtilTitle(location, x.tmplSheet, x.currentSheet)
	}

	if !location.isValid() {
		x.writeTitles(r.fields, titles)
	}

	if location.isValid() {
		x.rowsWritten = 0

		if r.isSlice {
			for i := 0; i < r.beanValue.Len(); i++ {
				x.writeTemplateRow(location, r.beanValue.Index(i), newSheet)
			}
		} else {
			x.writeTemplateRow(location, r.beanValue, newSheet)
		}

		x.removeTempleRows(location)
		x.mergeTitled(location, r.writeOption)

		return x.createTemplateDataValidations(location, x.currentSheet)
	}

	if r.isSlice {
		startRowNum := -1
		endRowNum := -1

		for i := 0; i < r.beanValue.Len(); i++ {
			rowNum := x.writeRow(r.fields, r.beanValue.Index(i))

			if i == 0 {
				startRowNum = int(rowNum)
			}

			endRowNum = int(rowNum)
		}
		x.mergeRows(r.fields, r.writeOption, startRowNum, endRowNum)
	} else {
		x.writeRow(r.fields, r.beanValue)
	}

	return x.createDataValidations(r.fields, x.currentSheet)
}

func copyRowsUtilTitle(location templateLocation, tmplSheet, dataSheet spreadsheet.Sheet) {
	for _, trow := range tmplSheet.Rows() {
		if trow.RowNumber() > location.titledRowNum {
			break
		}

		drow := dataSheet.Row(trow.RowNumber())
		copyRow(trow, drow)
	}
}

func copyRow(from, to spreadsheet.Row) {
	for _, f := range RowCells(from) {
		col, _ := f.Column()
		t := to.Cell(col)
		t.SetString(GetCellString(f))
		CopyCellStyle(f, t)
	}
}

func (x *Xlsx) writePlaceholder(fields []reflect.StructField,
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

// nolint:gomnd
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
	for _, tc := range l.titleFields {
		dv := tc.StructField.Tag.Get("dataValidation")
		if err := x.createColumnDataValidation(l.titledRowNum+1, sheet, dv, tc.Column); err != nil {
			return err
		}
	}

	return nil
}

func (x *Xlsx) createColumnDataValidation(startRowNum uint32, sheet spreadsheet.Sheet, tag, cellColumn string) error {
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
			return fmt.Errorf("unable to find sheet with name %s", dvSheetName) // nolint:goerr113
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

// Read reads the excel rows to slice.
// nolint:goerr113
func (x *Xlsx) Read(slicePtr interface{}) error {
	r := makeRun(slicePtr, nil)

	if !r.forRead() {
		return errors.New("the input argument should be a pointer of slice")
	}

	x.tmplSheet = x.createReadSheet(x.tmplWorkbook, r)
	x.currentSheet = x.createReadSheet(x.workbook, r)

	if r.asPlaceholder() {
		err := x.writePlaceholderToBean(r)
		if err != nil {
			return err
		}

		return nil
	}

	ignoreEmptyRows := r.ignoreEmptyRows()

	titles, customizedTitle := collectTitles(r.fields)
	loc, err := x.locateTitleRow(titles, customizedTitle)
	if err != nil {
		return err
	}

	location := *loc
	if location.isValid() {
		slice, err := x.readRows(r.beanType, location, ignoreEmptyRows)
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
			vs := GetCellString(x.currentSheet.Cell(v))
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

func (x *Xlsx) readRows(beanType reflect.Type, l templateLocation, ignoreEmptyRows bool) (reflect.Value, error) {
	slice := reflect.MakeSlice(reflect.SliceOf(beanType), 0, len(l.templateRows))

	for _, row := range l.templateRows {
		rowBean, err := x.createRowBean(beanType, l, row, ignoreEmptyRows)
		if err != nil {
			return reflect.Value{}, err
		}

		if rowBean.IsValid() {
			slice = reflect.Append(slice, rowBean)
		}
	}

	return slice, nil
}

func (x *Xlsx) createRowBean(beanType reflect.Type, l templateLocation,
	row spreadsheet.Row, ignoreEmptyRows bool) (reflect.Value, error) {
	type templateCellValue struct {
		TitleField
		value string
	}

	values := make([]templateCellValue, len(l.titleFields))
	emptyCells := 0

	for i, cell := range l.titleFields {
		c := row.Cell(cell.Column)
		s := GetCellString(c)
		values[i] = templateCellValue{
			TitleField: cell,
			value:      s,
		}

		if ignoreEmptyRows && s == "" {
			emptyCells++
		}
	}

	if emptyCells == len(l.titleFields) {
		return reflect.Value{}, nil
	}

	rowBean := reflect.New(beanType).Elem()

	for _, cell := range values {
		if err := setFieldValue(rowBean, cell.StructField, cell.value); err != nil {
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

	v, err := cast.ToAny(s, sf.Type)

	if err != nil && sf.Tag.Get("omiterr") != "true" {
		return err
	}

	f.Set(v)

	return nil
}

func (x *Xlsx) createWriteSheet(wb *spreadsheet.Workbook, r *run) (tmplSheet, dataSheet spreadsheet.Sheet) {
	wbSheet := spreadsheet.Sheet{}
	sheetName := r.FindTtag("sheet")

	if x.hasInput() {
		if sh := x.findSheet(wb, sheetName); sh.IsValid() {
			wbSheet = sh
		} else if len(wb.Sheets()) > 0 {
			wbSheet = wb.Sheets()[0]
		}
	}

	if !wbSheet.IsValid() {
		wbSheet = wb.AddSheet()
	}

	dataSheet = wbSheet

	if r.writeOption.SheetName != "" {
		if x.tmplSheetReused {
			dataSheet = wb.AddSheet()
		} else {
			x.tmplSheetReused = true
		}

		dataSheet.SetName(r.writeOption.SheetName)
	}

	if sheetName != "" && !strings.Contains(wbSheet.Name(), sheetName) {
		wbSheet.SetName(sheetName)
	}

	return wbSheet, dataSheet
}

func (x *Xlsx) createReadSheet(wb *spreadsheet.Workbook, r *run) spreadsheet.Sheet {
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
			return wb.Sheets()[0]
		}
	}

	return wbSheet
}

type TitleField struct {
	StructField reflect.StructField
	Title       string
	Column      string
}

func collectTitles(fields []reflect.StructField) ([]TitleField, bool) {
	titles := make([]TitleField, 0)
	customizedTitles := make([]TitleField, 0)

	for _, f := range fields {
		tf := TitleField{
			StructField: f,
			Title:       f.Name,
		}
		if t := f.Tag.Get("title"); t != "" {
			tf.Title = t
			customizedTitles = append(customizedTitles, tf)
			titles = append(titles, tf)
		} else {
			titles = append(titles, tf)
		}
	}

	if len(customizedTitles) > 0 {
		return customizedTitles, true
	}

	return titles, false
}

// SaveToFile writes the workbook out to a file.
func (x *Xlsx) SaveToFile(file string) error { return x.workbook.SaveToFile(file) }

// Save writes the workbook out to a writer in the zipped xlsx format.
func (x *Xlsx) Save(w io.Writer) error { return x.workbook.Save(w) }

func (x *Xlsx) writeRow(fields []reflect.StructField, value reflect.Value) uint32 {
	row := x.currentSheet.AddRow()
	x.rowsWritten++

	for _, field := range fields {
		setCellValue(row.AddCell(), field, value)
	}

	return row.RowNumber()
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

func (x *Xlsx) writeTitles(fields []reflect.StructField, titles []TitleField) {
	row := x.currentSheet.AddRow()

	for i := range fields {
		row.AddCell().SetString(titles[i].Title)
	}
}

type templateLocation struct {
	titledRowNum uint32
	titleFields  []TitleField
	templateRows []spreadsheet.Row
}

func (t *templateLocation) isValid() bool {
	return len(t.titleFields) > 0
}

var (
	ErrFailToLocationTitleRow = errors.New("unable to location title row")
	ErrNoExcelRead            = errors.New("no excel read")
)

func (x *Xlsx) locateTitleRow(titles []TitleField, customizedTitle bool) (*templateLocation, error) {
	if !x.hasInput() {
		return &templateLocation{}, ErrNoExcelRead
	}

	tmplSheet := x.tmplSheet
	if !tmplSheet.IsValid() {
		tmplSheet = x.currentSheet
	}

	rows := tmplSheet.Rows()
	titledRowNum, err := x.findTitledRow(titles, customizedTitle, rows)
	if err != nil {
		return nil, ErrFailToLocationTitleRow
	}

	templateRows := x.findTemplateRows(titledRowNum, rows)

	return &templateLocation{
		titledRowNum: titledRowNum,
		templateRows: templateRows,
		titleFields:  titles,
	}, nil
}

func (x *Xlsx) findTitledRow(titles []TitleField, customizedTitle bool, rows []spreadsheet.Row) (uint32, error) {
	for i, row := range rows {
		if i > 5 { // nolint:gomnd
			// 前5行都找不到的话，结束
			return 0, ErrFailToLocationTitleRow
		}

		found := false

		for _, cell := range RowCells(row) {
			cellString := GetCellString(cell)
			if cellString == "" {
				continue
			}

			for i, title := range titles {
				if !strings.Contains(cellString, title.Title) {
					continue
				}

				col, err := cell.Column()
				if err != nil {
					log.Printf("W! failed to get column error: %v", err)
					continue
				}

				if titles[i].Column != "" {
					return 0, perr.Wrapf(ErrFailToLocationTitleRow, "duplicate columns contains title %s", title.Title)
				}

				titles[i].Column = col
				found = true

				break
			}
		}

		if customizedTitle {
			found = func() bool {
				for _, t := range titles {
					if t.Column == "" {
						return false
					}
				}

				return true
			}()
		}

		if found {
			return row.RowNumber(), nil
		}
	}

	return 0, ErrFailToLocationTitleRow
}

func (x *Xlsx) findTemplateRows(titledRowNum uint32, rows []spreadsheet.Row) []spreadsheet.Row {
	templateRows := make([]spreadsheet.Row, 0)

	for _, row := range rows {
		if row.RowNumber() > titledRowNum {
			templateRows = append(templateRows, row)
		}
	}

	return templateRows
}

func (x *Xlsx) writeTemplateRow(l templateLocation, v reflect.Value, newSheet bool) {
	// 2 是为了计算row num(1-N), 从标题行(T)的下一行（T+1)开始写
	num := l.titledRowNum + 1 + x.rowsWritten
	x.rowsWritten++
	row := x.currentSheet.Row(num)

	for _, tc := range l.titleFields {
		setCellValue(row.Cell(tc.Column), tc.StructField, v)
	}

	x.copyRowStyle(l, row, newSheet)
}

func (x *Xlsx) copyRowStyle(l templateLocation, row spreadsheet.Row, newSheet bool) {
	if len(l.templateRows) == 0 {
		return
	}

	if !newSheet && x.rowsWritten < uint32(len(l.templateRows)) {
		return
	}

	templateRow := l.templateRows[(x.rowsWritten-1)%uint32(len(l.templateRows))]

	CopyRowStyle(templateRow, row)
}

func (x *Xlsx) removeTempleRows(l templateLocation) {
	if len(l.templateRows) == 0 {
		return // no template rows in the template file.
	}

	sheetData := x.currentSheet.X().CT_Worksheet.SheetData
	rows := sheetData.Row

	if endIndex := l.titledRowNum + x.rowsWritten; endIndex < uint32(len(rows)) {
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
		cellValue := GetCellString(x.currentSheet.Cell(k))

		if vars, ok := v.ParseVars(cellValue); ok {
			for vk, vv := range vars {
				plVars[vk] = vv
			}
		} else {
			log.Printf("W! failed to parse vars from cell value %s by Part %+v", cellValue, v.Content)
		}
	}

	return plVars
}

func (x *Xlsx) mergeTitled(l templateLocation, option WriteOption) {
	if x.rowsWritten <= 1 { // there is no need to mergeTitled for one row.
		return
	}

	switch option.MergeColsMode {
	case DoNotMerge:
		// nothing to do
	case MergeCols, MergeColsAlign:
		x.mergeCols(l, option.MergeColsMode)
	}
}

func (x *Xlsx) mergeCols(l templateLocation, mode MergeColsMode) {
	alignRowNums := make([]int, 0)

	for _, tf := range l.titleFields {
		var i uint32 = 1

		startCellString := ""
		startCell := spreadsheet.Cell{}
		startRowNum := 0
		lastCell := spreadsheet.Cell{}

		for ; i <= x.rowsWritten; i++ {
			rowNum := l.titledRowNum + i
			cell := x.currentSheet.Row(rowNum).Cell(tf.Column)
			cs := GetCellString(cell)
			if cs == startCellString && !reachAlignRowNum(alignRowNums, startRowNum, int(rowNum), mode) {
				lastCell = cell
				continue
			}

			if startCellString != "" && lastCell.X() != nil {
				x.currentSheet.AddMergedCells(startCell.Reference(), lastCell.Reference())

				alignRowNums = addAlignRowNum(alignRowNums, int(rowNum-1), mode)
			}

			startCellString = cs
			startCell = cell
			startRowNum = int(rowNum)
			lastCell = spreadsheet.Cell{}
		}

		if startCellString != "" && lastCell.X() != nil {
			x.currentSheet.AddMergedCells(startCell.Reference(), lastCell.Reference())
		}
	}
}

func (x *Xlsx) mergeRows(fields []reflect.StructField, option WriteOption, startRowNum, endRowNum int) {
	if endRowNum-startRowNum <= 1 {
		return
	}

	switch option.MergeColsMode {
	case DoNotMerge:
		// nothing to do
	case MergeCols, MergeColsAlign:
		x.mergeUntitledCols(fields, startRowNum, endRowNum, option.MergeColsMode)
	}
}

func (x *Xlsx) mergeUntitledCols(fields []reflect.StructField, startRow, endRow int, mode MergeColsMode) {
	alignRowNums := make([]int, 0)

	for i := range fields {
		startCellString := ""
		startCell := spreadsheet.Cell{}
		startRowNum := 0
		lastCell := spreadsheet.Cell{}

		for rowNum := startRow; rowNum <= endRow; rowNum++ {
			cell := x.currentSheet.Row(uint32(rowNum)).Cell(reference.IndexToColumn(uint32(i)))
			cs := GetCellString(cell)
			if cs == startCellString && !reachAlignRowNum(alignRowNums, startRowNum, rowNum, mode) {
				lastCell = cell
				continue
			}

			if startCellString != "" && lastCell.X() != nil {
				x.currentSheet.AddMergedCells(startCell.Reference(), lastCell.Reference())
				alignRowNums = addAlignRowNum(alignRowNums, rowNum-1, mode)
			}

			startCellString = cs
			startCell = cell
			startRowNum = rowNum
			lastCell = spreadsheet.Cell{}
		}

		if startCellString != "" && lastCell.X() != nil {
			x.currentSheet.AddMergedCells(startCell.Reference(), lastCell.Reference())
		}
	}
}

func addAlignRowNum(alignRowNums []int, num int, mode MergeColsMode) []int {
	if mode == MergeColsAlign {
		alignRowNums = append(alignRowNums, num)
		sort.Ints(alignRowNums)
	}

	return alignRowNums
}

func reachAlignRowNum(alignRowNums []int, startRowNum, num int, mode MergeColsMode) bool {
	if mode == MergeCols {
		return false
	}

	for _, align := range alignRowNums {
		if startRowNum <= align && align <= num {
			return true
		}
	}

	return false
}

// nolint:gochecknoglobals
var (
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
	if s == "" {
		return time.Time{}, nil
	}

	if f := tag.Get("format"); f != "" {
		return time.ParseInLocation(ParseJavaTimeFormat(f), s, time.Local)
	}

	return dateparse.ParseLocal(s)
}
