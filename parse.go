package excelizemapper

import (
	"fmt"
	"log/slog"
	"reflect"
	"sort"
	"strconv"
	"strings"
)

type parser struct {
	tagKey   string
	autosort bool

	tagDelim         string
	tagHeaderKey     string
	tagIndexKey      string
	tagDefaultKey    string
	tagFormatKey     string
	tagWidthKey      string
	tagDynamicKey    string
	tagDynamicPosKey string
	tagDynamicValKey string
}

func (p *parser) parse(data interface{}) ([]Column, *DynamicRules, error) {
	dv := reflect.ValueOf(data)
	di := reflect.Indirect(dv)
	dk := di.Kind()

	if dk != reflect.Array && dk != reflect.Slice {
		return nil, nil, fmt.Errorf("data not array or slice")
	}

	itemType := di.Type().Elem()
	if itemType.Kind() == reflect.Ptr {
		itemType = itemType.Elem()
	}

	cols, rules, err := p.parseFieldsRecursive(itemType, "")
	if err != nil {
		return nil, nil, err
	}

	sort.Slice(cols, func(i, j int) bool {
		return cols[i].ColumnIndex < cols[j].ColumnIndex
	})

	return cols, rules, nil
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

type DynamicRules struct {
	Mappings        map[string]string
	ValueField      string
	ParentFieldName string
	ParentRule      string
}

func (dr *DynamicRules) getReplacedHeader(entryVal reflect.Value) string {
	colHeader := dr.ParentRule
	for pos, field := range dr.Mappings {
		slog.Debug("field", "name", field)

		fieldRef := entryVal.FieldByName(field)
		slog.Debug("fieldRef", "name", fieldRef.Interface())

		colHeader = strings.ReplaceAll(colHeader, pos, fmt.Sprint(fieldRef.Interface()))
		slog.Debug("colHeader", "name", colHeader)
	}
	return colHeader
}

func (p *parser) getDynamicRules(dynamicSlice reflect.StructField) *DynamicRules {

	mappings := make(map[string]string)
	var valField string

	t := dynamicSlice.Type.Elem()
	for i := 0; i < t.NumField(); i++ {

		field := t.Field(i)

		if !field.IsExported() {
			continue
		}

		tags := p.getTagsByKey(field)
		slog.Debug("parsed tags", "value", tags)

		if tags == nil {
			continue
		}

		if posKey, ok := tags[p.tagDynamicPosKey]; ok {
			mappings[posKey] = field.Name
			continue
		}

		if _, ok := tags[p.tagDynamicValKey]; ok {
			valField = field.Name
			slog.Debug("found valField", "value", valField)
			continue
		}

	}

	return &DynamicRules{
		Mappings:        mappings,
		ValueField:      valField,
		ParentFieldName: dynamicSlice.Name,
	}

}

func (p *parser) getTagsByKey(dynamicSlice reflect.StructField) map[string]string {
	fullTagVal := dynamicSlice.Tag.Get(p.tagKey)
	if fullTagVal == "" {
		return nil
	}

	return p.parseTags(fullTagVal)
}

func (p *parser) parseFieldsRecursive(t reflect.Type, prefix string) ([]Column, *DynamicRules, error) {
	var cols []Column
	var dynamicRules *DynamicRules
	autoIndex := 0

	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)

		if !field.IsExported() {
			continue
		}

		if field.Type.Kind() == reflect.Struct && field.Anonymous {
			nestedCols, _, err := p.parseFieldsRecursive(field.Type, prefix+field.Name+".")
			if err != nil {
				return nil, nil, err
			}
			cols = append(cols, nestedCols...)
			continue
		}

		tags := p.getTagsByKey(field)
		if tags == nil {
			continue
		}

		parentRule, hasDynamicTag := tags[p.tagDynamicKey]
		if field.Type.Kind() == reflect.Slice && hasDynamicTag {
			dynamicRules = p.getDynamicRules(field)
			dynamicRules.ParentRule = parentRule
			continue
		}

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
				return nil, nil, fmt.Errorf("invalid index value %q for field %s", indexStr, field.Name)
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

	return cols, dynamicRules, nil
}
