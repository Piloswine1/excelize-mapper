package excelizemapper

import (
	"fmt"
	"reflect"

	"github.com/xuri/excelize/v2"
)

const defaultTagKey = "excelize-mapper"

type ExcelizeMapper struct {
	ops options
}

func NewExcelizeMapper(opts ...Option) ExcelizeMapper {
	op := options{
		TagKey: defaultTagKey,
	}

	for _, opt := range opts {
		opt(&op)
	}

	return ExcelizeMapper{
		ops: op,
	}
}

func (em *ExcelizeMapper) SetData(f *excelize.File, sheet string, slice interface{}) error {
	p := parser{
		tagKey:        em.ops.TagKey,
		tagDelim:      ";",
		tagHeaderKey:  "header",
		tagIndexKey:   "index",
		tagDefaultKey: "default",
		tagFormatKey:  "format",
		tagWidthKey:   "width",
	}

	cells, err := p.parse(slice)
	if err != nil {
		return err
	}

	headers := make([]string, 0, len(cells))
	for _, cell := range cells {
		headers = append(headers, cell.HeaderName)
	}

	err = f.SetSheetRow(sheet, "A1", &headers)
	if err != nil {
		return fmt.Errorf("excelize SetSheetRow error: %w", err)
	}

	di := reflect.Indirect(reflect.ValueOf(slice))
	for rowIndex := 0; rowIndex < di.Len(); rowIndex++ {
		rowVal := reflect.Indirect(di.Index(rowIndex))
		vals := make([]interface{}, 0, len(cells))
		for _, fieldInfo := range cells {
			fieldValue := rowVal.FieldByName(fieldInfo.FieldName)

			if fieldValue.IsZero() && fieldInfo.DefaultValue != "" {
				fieldValue = reflect.ValueOf(fieldInfo.DefaultValue)
			}

			if format, ok := em.ops.FormatterMap[fieldInfo.FormatterKey]; ok {
				formatVal := format(fieldValue.Interface())
				fieldValue = reflect.ValueOf(formatVal)
			}

			vals = append(vals, fieldValue.Interface())
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
