package excelizemapper

import (
	"fmt"
	"reflect"

	"github.com/xuri/excelize/v2"
)

const (
	defaultTagKey        = "excelize-mapper"
	defaultTagDelim      = ";"
	defaultTagHeaderKey  = "header"
	defaultTagIndexKey   = "index"
	defaultTagWidthKey   = "width"
	defaultTagFormatKey  = "format"
	defaultTagDefaultKey = "default"
)

type ExcelizeMapper struct {
	options options
	parser  parser
}

func NewExcelizeMapper(opts ...Option) ExcelizeMapper {
	op := options{
		tagKey:       defaultTagKey,
		autoSort:     true,
		formatterMap: make(map[string]Format, 0),
	}

	for _, opt := range opts {
		opt(&op)
	}

	return ExcelizeMapper{
		options: op,
		parser: parser{
			tagKey:        op.tagKey,
			autosort:      op.autoSort,
			tagDelim:      defaultTagDelim,
			tagHeaderKey:  defaultTagHeaderKey,
			tagIndexKey:   defaultTagIndexKey,
			tagDefaultKey: defaultTagDefaultKey,
			tagFormatKey:  defaultTagFormatKey,
			tagWidthKey:   defaultTagWidthKey,
		},
	}
}

func (em *ExcelizeMapper) SetData(f *excelize.File, sheet string, slice interface{}) error {
	columns, err := em.parser.parse(slice)
	if err != nil {
		return err
	}

	headers := make([]string, 0, len(columns))
	currentIndex := 0
	for _, column := range columns {
		for ; currentIndex < column.ColumnIndex; currentIndex++ { // skip index
			headers = append(headers, "")
		}

		headers = append(headers, column.HeaderName)
		currentIndex++

		width := em.options.defaultWidth
		if column.ColumnWidth > 0 {
			width = column.ColumnWidth
		}
		if width > 0 { // set column width
			col, err := excelize.ColumnNumberToName(currentIndex)
			if err != nil {
				return fmt.Errorf("excelize ColumnNumberToName error: %w", err)
			}
			f.SetColWidth(sheet, col, col, width)
		}
	}

	err = f.SetSheetRow(sheet, "A1", &headers)
	if err != nil {
		return fmt.Errorf("excelize SetSheetRow error: %w", err)
	}

	di := reflect.Indirect(reflect.ValueOf(slice))
	for rowIndex := 0; rowIndex < di.Len(); rowIndex++ {
		rowVal := reflect.Indirect(di.Index(rowIndex))
		vals := make([]interface{}, 0, len(columns))

		currentIndex := 0
		for _, column := range columns {
			for ; currentIndex < column.ColumnIndex; currentIndex++ { // skip index
				vals = append(vals, "")
			}
			fieldValue := rowVal.FieldByName(column.FieldName)

			if fieldValue.Kind() == reflect.Ptr {
				if fieldValue.IsNil() {
					fieldValue = (reflect.Zero(fieldValue.Type()))
				} else {
					fieldValue = (fieldValue.Elem())
				}
			} else if fieldValue.IsZero() && column.DefaultValue != "" {
				fieldValue = reflect.ValueOf(column.DefaultValue)
			} else {
				fieldValue = (fieldValue)
			}

			if format, ok := em.options.formatterMap[column.FormatterKey]; ok {
				formatVal := format(fieldValue.Interface())
				fieldValue = reflect.ValueOf(formatVal)
			}

			vals = append(vals, fieldValue.Interface())
			currentIndex++
		}

		cell, err := excelize.CoordinatesToCellName(1, rowIndex+2)
		if err != nil {
			return fmt.Errorf("excelize CoordinatesToCellName error: %w", err)
		}
		err = f.SetSheetRow(sheet, cell, &vals)
		if err != nil {
			return fmt.Errorf("excelize SetSheetRow error: %w", err)
		}
	}

	return nil
}
