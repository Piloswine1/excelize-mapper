package excelizemapper

import (
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

type baseModel struct {
	Int   int   `excelize-mapper:"header:Int"`
	Int8  int8  `excelize-mapper:"header:Int8"`
	Int16 int16 `excelize-mapper:"header:Int16"`
	Int32 int32 `excelize-mapper:"header:Int32"`
}

var baseData = baseModel{
	Int:   int(1<<31 - 1),
	Int8:  int8(1<<7 - 1),
	Int16: int16(1<<15 - 1),
	Int32: int32(1<<31 - 1),
}

type customSortModel struct {
	Int   int   `excelize-mapper:"index:0;header:Int"`
	Int8  int8  `excelize-mapper:"index:1;header:Int8"`
	Int16 int16 `excelize-mapper:"index:2;header:Int16"`
	Int32 int32 `excelize-mapper:"index:3;header:Int32"`
	Int64 int64 `excelize-mapper:"index:4;header:Int64"`

	Uint   uint   `excelize-mapper:"index:5;header:Uint"`
	Uint8  uint8  `excelize-mapper:"index:6;header:Uint8"`
	Uint16 uint16 `excelize-mapper:"index:7;header:Uint16"`
	Uint32 uint32 `excelize-mapper:"index:8;header:Uint32"`
	Uint64 uint64 `excelize-mapper:"index:9;header:Uint64"`

	Float32 float32 `excelize-mapper:"index:10;header:Float32"`
	Float64 float64 `excelize-mapper:"index:11;header:Float64"`

	Byte   byte   `excelize-mapper:"index:12;header:Byte"`
	Rune   rune   `excelize-mapper:"index:13;header:Rune"`
	String string `excelize-mapper:"index:14;header:String"`
	Bool   bool   `excelize-mapper:"index:15;header:Bool"`

	Time      time.Time `excelize-mapper:"index:16;header:Time"`
	NextIndex string    `excelize-mapper:"index:18;header:NextIndex"` // skip index 17
}

var customSortData = customSortModel{
	Int:    int(1<<31 - 1),
	Int8:   int8(1<<7 - 1),
	Int16:  int16(1<<15 - 1),
	Int32:  int32(1<<31 - 1),
	Int64:  int64(1<<63 - 1),
	Uint:   uint(1<<32 - 1),
	Uint8:  uint8(1<<8 - 1),
	Uint16: uint16(1<<16 - 1),
	Uint32: uint32(1<<32 - 1),
	Uint64: uint64(1<<63 - 1),

	Float32: float32(100.1234),
	Float64: float64(100.1234),

	Byte:   byte(1<<8 - 1),
	Rune:   rune(1<<31 - 1),
	String: "string",
	Bool:   true,

	Time:      time.Now(),
	NextIndex: "nextIndex",
}

func TestSetData(t *testing.T) {
	sheetName := "sheet1"

	originData := make([]baseModel, 0)
	originData = append(originData, baseData, baseData)

	f := excelize.NewFile()
	defer f.Close()

	mapper := NewExcelizeMapper()

	err := mapper.SetData(f, sheetName, originData)
	if err != nil {
		t.Fatal(err)
	}

	f.SaveAs("./testData/base.xlsx")
}

func TestCustomSortSetData(t *testing.T) {
	sheetName := "sheet1"

	originData := make([]customSortModel, 0)
	originData = append(originData, customSortData, customSortData)

	f := excelize.NewFile()
	defer f.Close()

	mapper := NewExcelizeMapper(WithAutoSort(false))

	err := mapper.SetData(f, sheetName, originData)
	if err != nil {
		t.Fatal(err)
	}

	f.SaveAs("./testData/custom_sort.xlsx")
}

type targetValueModel struct {
	Disable  bool   `excelize-mapper:"header:Disable;format:disable"`
	NotEmpty string `excelize-mapper:"header:NotEmpty;default:use default value"`
}

func TestCustomValueSetData(t *testing.T) {
	sheetName := "sheet1"

	originData := make([]targetValueModel, 0)
	originData = append(originData,
		targetValueModel{Disable: true, NotEmpty: ""},
		targetValueModel{Disable: false, NotEmpty: "not empty"},
	)

	f := excelize.NewFile()
	defer f.Close()

	disableFormat := func(v interface{}) string {
		if d, ok := v.(bool); ok {
			if d {
				return "enabled"
			}
			return "disabled"
		}
		return ""
	}

	mapper := NewExcelizeMapper(
		WithFormatter("disable", disableFormat),
	)

	err := mapper.SetData(f, sheetName, originData)
	if err != nil {
		t.Fatal(err)
	}

	f.SaveAs("./testData/custom_value.xlsx")
}

type styleModel struct {
	LongText string `excelize-mapper:"header:LongText;width:50"`
}

func TestCustomStyleSetData(t *testing.T) {
	sheetName := "sheet1"

	originData := make([]styleModel, 0)
	originData = append(originData,
		styleModel{LongText: "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"},
		styleModel{LongText: "bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb"},
	)

	mapper := NewExcelizeMapper()

	f := excelize.NewFile()
	defer f.Close()
	err := mapper.SetData(f, sheetName, originData)
	if err != nil {
		t.Fatal(err)
	}

	f.SaveAs("./testData/custom_style.xlsx")
}
