package main

import (
	"strconv"
	"strings"
	"time"

	excelizemapper "github.com/Piloswine1/excelize-mapper"

	"github.com/xuri/excelize/v2"
)

type ExcelExportOptions struct {
	AutoSort     bool
	DefaultWidth float64
	FormatterMap map[string]excelizemapper.Format
}

type ExcelExportOption func(*ExcelExportOptions)

func WithAutoSort(autoSort bool) ExcelExportOption {
	return func(o *ExcelExportOptions) {
		o.AutoSort = autoSort
	}
}

func WithDefaultWidth(width float64) ExcelExportOption {
	return func(o *ExcelExportOptions) {
		o.DefaultWidth = width
	}
}

func WithFormatter(name string, f excelizemapper.Format) ExcelExportOption {
	return func(o *ExcelExportOptions) {
		if o.FormatterMap == nil {
			o.FormatterMap = make(map[string]excelizemapper.Format)
		}
		o.FormatterMap[name] = f
	}
}

func main() {
	example1(
		WithAutoSort(false),
		WithFormatter("slice", excelizemapper.SliceFormatter),
		WithDefaultWidth(70),
		WithFormatter("sex", SexFormat),
	)
	example2()
}

type IdInt struct {
	Id int `excelize-mapper:"header:Номер;width:50;index:0;"`
}

type Sex int32

const (
	SexMale Sex = iota
	SexFemale
)

type User struct {
	IdInt
	Name      string    `excelize-mapper:"header:Name;index:4;"`
	Desc      *string   `excelize-mapper:"header:Desc;width:50;index:3;"`
	Sex       Sex       `excelize-mapper:"header:Sex;format:sex;index:2;"`
	Address   string    `excelize-mapper:"header:Address;default:China;index:1;"`
	CreatedAt time.Time `excelize-mapper:"header:CreatedAt;index:5;"`
}

var SexFormat = func(value interface{}) string {
	switch value.(Sex) {
	case SexMale:
		return "Male"
	case SexFemale:
		return "Female"
	default:
		return "Unknown"
	}
}

const dt = "2023-12-21T15:38:29.808+08:00"

func example1(opts ...ExcelExportOption) {
	options := ExcelExportOptions{
		AutoSort:     true,
		DefaultWidth: 50,
		FormatterMap: make(map[string]excelizemapper.Format),
	}

	for _, opt := range opts {
		opt(&options)
	}

	ct, _ := time.Parse(time.RFC3339, dt)
	data := []User{
		{
			IdInt:     IdInt{Id: 1},
			Name:      "Tom",
			Sex:       SexMale,
			Address:   "Singapore",
			CreatedAt: ct,
		}, {
			IdInt:     IdInt{Id: 2},
			Name:      "Jerry",
			Sex:       SexFemale,
			Address:   "",
			CreatedAt: ct,
		},
	}
	sstr := "qwerty qaz"
	data[0].Desc = &sstr

	f := excelize.NewFile()
	defer f.Close()

	m := excelizemapper.NewExcelizeMapper(
		excelizemapper.WithAutoSort(options.AutoSort),
		excelizemapper.WithDefaultWidth(options.DefaultWidth),
		excelizemapper.WithFormatter("slice", excelizemapper.SliceFormatter),
		excelizemapper.WithFormatter("sex", SexFormat),
	)

	for name, formatter := range options.FormatterMap {
		m.SetFormatter(name, formatter)
	}

	err := m.SetData(f, "Sheet1", data)
	if err != nil {
		panic(err)
	}

	err = f.SaveAs("./example1.xlsx")
	if err != nil {
		panic(err)
	}
}

type User2 struct {
	IdInt
	Name string    `excelize-mapper:"header:Name;"`
	Desc string    `excelize-mapper:"header:Desc;"`
	Sex  Sex       `excelize-mapper:"header:Sex;format:sex;"`
	Arr  []int     `excelize-mapper:"header:Arr;format:slice;"`
	Time time.Time `excelize-mapper:"header:Time;format:time;"`
}

type MyArr []int

func (m MyArr) String() string {
	builder := strings.Builder{}

	for i, v := range m {
		if i > 0 {
			builder.WriteString(", ")
		}
		val := strconv.Itoa(v)
		builder.WriteString(val)
	}

	return builder.String()
}

// custom index
func example2() {
	data := []User2{{
		IdInt: IdInt{Id: 1},
		Name:  "Tom",
		Desc:  "This is a long text, it will be wrapped.",
		Sex:   SexMale,
		Arr:   []int{1, 23, 45},
		Time:  time.Now(),
	}, {
		IdInt: IdInt{Id: 2},
		Name:  "Jerry",
		Desc:  "This is a long text.",
		Sex:   SexFemale,
		Arr:   []int{1, 23, 0},
		Time:  time.Now(),
	}}

	f := excelize.NewFile()
	defer f.Close()

	m := excelizemapper.NewExcelizeMapper(
		excelizemapper.WithAutoSort(true),
		excelizemapper.WithDefaultWidth(40),
		excelizemapper.WithFormatter("sex", SexFormat),
		excelizemapper.WithFormatter("slice", excelizemapper.SliceFormatter),
	)
	err := m.SetData(f, "Sheet1", data)
	if err != nil {
		panic(err)
	}

	err = f.SaveAs("./example2.xlsx")
	if err != nil {
		panic(err)
	}
}
