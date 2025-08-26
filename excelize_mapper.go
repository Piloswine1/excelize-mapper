package excelizemapper

import (
	"fmt"
	"log/slog"
	"reflect"
	"slices"
	"strings"

	"github.com/xuri/excelize/v2"
)

const (
	defaultTagKey           = "excelize-mapper"
	defaultTagDelim         = ";"
	defaultTagHeaderKey     = "header"
	defaultTagIndexKey      = "index"
	defaultTagWidthKey      = "width"
	defaultTagFormatKey     = "format"
	defaultTagDefaultKey    = "default"
	defaultTagDynamicKey    = "dynamic"
	defaultTagDynamicPosKey = "dynamicpos"
	defaultTagDynamicValKey = "dynamicval"
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
			tagKey:           op.tagKey,
			autosort:         op.autoSort,
			tagDelim:         defaultTagDelim,
			tagHeaderKey:     defaultTagHeaderKey,
			tagIndexKey:      defaultTagIndexKey,
			tagDefaultKey:    defaultTagDefaultKey,
			tagFormatKey:     defaultTagFormatKey,
			tagWidthKey:      defaultTagWidthKey,
			tagDynamicKey:    defaultTagDynamicKey,
			tagDynamicPosKey: defaultTagDynamicPosKey,
			tagDynamicValKey: defaultTagDynamicValKey,
		},
	}
}

// Already know data is slice
func (em *ExcelizeMapper) parseSlice(rules *DynamicRules, model interface{}) []string {
	var headers []string
	modelValue := reflect.ValueOf(model)
	for i := 0; i < modelValue.Len(); i++ {

		modelEntry := modelValue.Index(i)
		sliceEntries := modelEntry.FieldByName(rules.ParentFieldName)

		for j := 0; j < sliceEntries.Len(); j++ {
			entryVal := sliceEntries.Index(j)

			header := rules.getReplacedHeader(entryVal)

			if !slices.Contains(headers, header) {
				headers = append(headers, header)
			}
		}

	}

	return headers
}

func (em *ExcelizeMapper) foreachValues(rules *DynamicRules, modelValue reflect.Value, cb func(string, string)) {

	sliceEntries := modelValue.FieldByName(rules.ParentFieldName)
	slog.Debug("modelValue",
		"kind", sliceEntries.Type().Kind().String(),
		"name", modelValue.Type().Name())

	for j := 0; j < sliceEntries.Len(); j++ {
		entry := sliceEntries.Index(j)
		slog.Debug("entryVal", "name", entry.Type().Name())

		header := rules.getReplacedHeader(entry)
		slog.Debug("getReplacedHeader", "name", header)
		slog.Debug("ValueField", "name", rules.ValueField)

		val := entry.FieldByName(rules.ValueField)
		slog.Debug("val", "value", val)

		cb(header, fmt.Sprint(val))

	}
}

func (em *ExcelizeMapper) SetData(f *excelize.File, sheet string, slice interface{}) error {
	columns, dynamicRules, err := em.parser.parse(slice)
	if err != nil {
		return err
	}

	headers := make([]string, 0, len(columns))
	currentIndex := 0
	for _, column := range columns {
		for ; currentIndex < column.ColumnIndex; currentIndex++ {
			headers = append(headers, "")
		}

		headers = append(headers, column.HeaderName)
		currentIndex = column.ColumnIndex + 1
	}

	for _, column := range columns {
		width := em.options.defaultWidth
		if column.ColumnWidth > 0 {
			width = column.ColumnWidth
		}
		if width > 0 {
			colName, err := excelize.ColumnNumberToName(column.ColumnIndex + 1)
			if err != nil {
				return fmt.Errorf("excelize ColumnNumberToName error: %w", err)
			}
			f.SetColWidth(sheet, colName, colName, width)
		}
	}

	// Handle dynamic fields headers
	var dynamicHeaders []string

	if dynamicRules != nil {
		dynamicHeaders = em.parseSlice(dynamicRules, slice)
		if dynamicHeaders != nil {
			headers = append(headers, dynamicHeaders...)
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

		currentIndex = 0
		for _, column := range columns {
			for ; currentIndex < column.ColumnIndex; currentIndex++ {
				vals = append(vals, "")
			}

			fieldValue := getNestedFieldValue(rowVal, column.FieldName)

			if fieldValue.Kind() == reflect.Ptr {
				if fieldValue.IsNil() {
					fieldValue = reflect.Zero(fieldValue.Type().Elem())
				} else {
					fieldValue = fieldValue.Elem()
				}
			} else if fieldValue.IsZero() && column.DefaultValue != "" {
				fieldValue = reflect.ValueOf(column.DefaultValue)
			}

			if format, ok := em.options.formatterMap[column.FormatterKey]; ok {
				formatVal := format(fieldValue.Interface())
				fieldValue = reflect.ValueOf(formatVal)
			}

			vals = append(vals, fieldValue.Interface())

			currentIndex = column.ColumnIndex + 1
		}

		// Handle dynamic fields values
		if len(dynamicHeaders) > 0 {
			dynamicVals := make([]interface{}, len(dynamicHeaders))

			em.foreachValues(dynamicRules, rowVal, func(niddle, val string) {
				pos := slices.IndexFunc(dynamicHeaders, func(header string) bool {
					return header == niddle
				})
				dynamicVals[pos] = val
			})

			vals = append(vals, dynamicVals...)
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

func getNestedFieldValue(v reflect.Value, fieldPath string) reflect.Value {
	parts := strings.Split(fieldPath, ".")
	for _, part := range parts {
		if v.Kind() == reflect.Ptr {
			if v.IsNil() {
				v = reflect.Zero(v.Type().Elem())
			} else {
				v = v.Elem()
			}
		}
		v = v.FieldByName(part)
	}
	return v
}
