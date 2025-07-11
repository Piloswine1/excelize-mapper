package excelizemapper

import (
	"fmt"
	"reflect"
	"sort"
	"strconv"
	"strings"
)

type parser struct {
	tagKey   string
	autosort bool

	tagDelim      string
	tagHeaderKey  string
	tagIndexKey   string
	tagDefaultKey string
	tagFormatKey  string
	tagWidthKey   string
}

func (p *parser) parse(data interface{}) ([]Column, error) {
	dv := reflect.ValueOf(data)
	di := reflect.Indirect(dv)
	dk := di.Kind()

	if dk != reflect.Array && dk != reflect.Slice {
		return nil, fmt.Errorf("data not array or slice")
	}

	itemType := di.Type().Elem()
	if itemType.Kind() == reflect.Ptr {
		itemType = itemType.Elem()
	}

	cols, err := p.parseFieldsRecursive(itemType, "")
	if err != nil {
		return nil, err
	}

	sort.Slice(cols, func(i, j int) bool {
		return cols[i].ColumnIndex < cols[j].ColumnIndex
	})

	return cols, nil
}

func (p *parser) parseTags(tag string) map[string]string {
	kv := make(map[string]string)
	tags := strings.Split(tag, p.tagDelim)
	for _, t := range tags {
		t = strings.TrimSpace(t)
		if t == "" {
			continue
		}

		kvSlice := strings.SplitN(t, ":", 2)
		if len(kvSlice) != 2 {
			continue
		}

		kv[kvSlice[0]] = kvSlice[1]
	}
	return kv
}

func (p *parser) parseFieldsRecursive(t reflect.Type, prefix string) ([]Column, error) {
	var cols []Column
	autoIndex := 0

	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)

		if !field.IsExported() {
			continue
		}

		if field.Type.Kind() == reflect.Struct && field.Anonymous {
			nestedCols, err := p.parseFieldsRecursive(field.Type, prefix+field.Name+".")
			if err != nil {
				return nil, err
			}
			cols = append(cols, nestedCols...)
			continue
		}

		fullTagVal := field.Tag.Get(p.tagKey)
		if fullTagVal == "" {
			continue
		}

		tags := p.parseTags(fullTagVal)
		header, ok := tags[p.tagHeaderKey]
		if !ok {
			continue
		}

		var colWidth float64
		if widthStr, ok := tags[p.tagWidthKey]; ok {
			if val, err := strconv.ParseFloat(widthStr, 64); err == nil {
				colWidth = val
			}
		}

		var colIndex int
		if !p.autosort {
			indexStr, ok := tags[p.tagIndexKey]
			if !ok {
				continue
			}
			idx, err := strconv.Atoi(indexStr)
			if err != nil {
				return nil, fmt.Errorf("invalid index value %q for field %s", indexStr, field.Name)
			}
			colIndex = idx
		} else {
			colIndex = autoIndex
			autoIndex++
		}

		col := Column{
			ColumnIndex:  colIndex,
			HeaderName:   header,
			ColumnWidth:  colWidth,
			DefaultValue: tags[p.tagDefaultKey],
			FormatterKey: tags[p.tagFormatKey],
			FieldName:    prefix + field.Name,
		}

		cols = append(cols, col)
	}

	return cols, nil
}
