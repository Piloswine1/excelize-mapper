package excelizemapper

import (
	"fmt"
	"reflect"
	"strings"
)

/*
Return array in format: "1, 2, 3, ...".

	 How to use:
	 In struct use excelize-mapper:"format:slice;"
	 and
	 m := excelizemapper.NewExcelizeMapper(
			excelizemapper.WithFormatter("slice", excelizemapper.SliceFormatter),
		)
*/
func SliceFormatter(val interface{}) string {
	v := reflect.ValueOf(val)
	if v.Kind() != reflect.Slice && v.Kind() != reflect.Array {
		return ""
	}

	parts := make([]string, v.Len())
	for i := 0; i < v.Len(); i++ {
		parts[i] = fmt.Sprintf("%v", v.Index(i).Interface())
	}
	return strings.Join(parts, ", ")
}

func (em *ExcelizeMapper) SetFormatter(name string, format Format) {
	em.options.formatterMap[name] = format
}
