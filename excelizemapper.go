package excelizemapper

import (
	"fmt"
	"reflect"

	"github.com/xuri/excelize/v2"
)

const defaultTagKey = "excelize-mapper"

type ExcelizeMapper struct {
	options options
	parser  parser
}

func NewExcelizeMapper(opts ...Option) ExcelizeMapper {
	op := options{
		TagKey: defaultTagKey,
	}

	for _, opt := range opts {
		opt(&op)
	}

	return ExcelizeMapper{
		options: op,
		parser: parser{
			tagKey:        op.TagKey,
			tagDelim:      ";",
			tagHeaderKey:  "header",
			tagIndexKey:   "index",
			tagDefaultKey: "default",
			tagFormatKey:  "format",
			tagWidthKey:   "width",
		},
	}
}

func (em *ExcelizeMapper) SetData(f *excelize.File, sheet string, slice interface{}) error {
	cells, err := em.parser.parse(slice)
	if err != nil {
		return err
	}

	headers := make([]string, 0, len(cells))
	currentIndex := 0
	for _, cell := range cells {
		for ; currentIndex < cell.CellIndex; currentIndex++ { // skip index
			headers = append(headers, "")
		}

		headers = append(headers, cell.HeaderName)
		currentIndex++
	}

	err = f.SetSheetRow(sheet, "A1", &headers)
	if err != nil {
		return fmt.Errorf("excelize SetSheetRow error: %w", err)
	}

	di := reflect.Indirect(reflect.ValueOf(slice))
	for rowIndex := 0; rowIndex < di.Len(); rowIndex++ {
		rowVal := reflect.Indirect(di.Index(rowIndex))
		vals := make([]interface{}, 0, len(cells))
		currentIndex := 0
		for _, cell := range cells {
			for ; currentIndex < cell.CellIndex; currentIndex++ { // skip index
				vals = append(vals, "")
			}
			fieldValue := rowVal.FieldByName(cell.FieldName)

			if fieldValue.IsZero() && cell.DefaultValue != "" {
				fieldValue = reflect.ValueOf(cell.DefaultValue)
			}

			if format, ok := em.options.FormatterMap[cell.FormatterKey]; ok {
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

// func (em *ExcelizeMapper) GetData(f *excelize.File, sheet string, slice interface{}) error {

// }
