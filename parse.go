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

func (p *parser) parse(data interface{}) ([]Cell, error) {
	dv := reflect.ValueOf(data)
	di := reflect.Indirect(dv)
	dk := di.Kind()

	if dk != reflect.Array && dk != reflect.Slice {
		return nil, fmt.Errorf("datas not array or slice")
	}

	itemType := di.Type().Elem()
	if itemType.Kind() == reflect.Ptr {
		itemType = itemType.Elem()

	}

	cells := make([]Cell, 0, 10)
	index := 0
	for fieldIndex := 0; fieldIndex < itemType.NumField(); fieldIndex++ {
		fullTagVal := itemType.Field(fieldIndex).Tag.Get(p.tagKey)

		if fullTagVal != "" {
			tags := p.parseTags(fullTagVal)
			if _, ok := tags[p.tagHeaderKey]; ok {
				var cellWidth float64
				if widthStr, ok := tags[p.tagWidthKey]; ok {
					if val, err := strconv.ParseFloat(widthStr, 64); err == nil {
						cellWidth = val
					}
				}

				if !p.autosort { // use custom index
					cellIndex, ok := tags[p.tagIndexKey]
					if !ok {
						continue
					}
					index, _ = strconv.Atoi(cellIndex)
				}

				cell := Cell{
					CellIndex:    index,
					HeaderName:   tags[p.tagHeaderKey],
					CellWidth:    cellWidth,
					DefaultValue: tags[p.tagDefaultKey],
					FormatterKey: tags[p.tagFormatKey],
					FieldName:    itemType.Field(fieldIndex).Name,
				}
				cells = append(cells, cell)
			}
			index++
		}
	}

	sort.Slice(cells, func(i, j int) bool {
		return cells[i].CellIndex < cells[j].CellIndex
	})

	return cells, nil
}

func (p *parser) parseTags(tag string) map[string]string {
	kv := make(map[string]string)
	tags := strings.Split(tag, p.tagDelim)
	for _, t := range tags {
		t = strings.TrimSpace(t)
		if t == "" {
			continue
		}

		kvSlice := strings.Split(t, ":")
		if len(kvSlice) != 2 {
			continue
		}

		kv[kvSlice[0]] = kvSlice[1]
	}
	return kv
}
