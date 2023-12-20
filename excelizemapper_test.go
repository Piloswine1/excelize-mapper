package excelizemapper

import (
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

type baseExportModel struct {
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

	Time time.Time `excelize-mapper:"index:16;header:Time"`
}

var baseExportData = baseExportModel{
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

	Time: time.Now(),
}

func TestBaseTypeExportExcel(t *testing.T) {
	sheetName := "sheet1"

	originData := make([]baseExportModel, 0)
	originData = append(originData, baseExportData, baseExportData)

	f := excelize.NewFile()
	defer f.Close()

	mapper := NewExcelizeMapper()

	err := mapper.SetData(f, sheetName, originData)
	if err != nil {
		t.Fatal(err)
	}

	f.SaveAs("./testData/test.xlsx")
}

// func TestSlicePtrStructExportExcel(t *testing.T) {
// 	sheetName := "sheet1"

// 	originData := make([]*baseExportModel, 0)
// 	originData = append(originData, &baseExportData, &baseExportData)

// 	file := xlsx.NewFile()
// 	err := ExportExcel(file, sheetName, originData)
// 	if err != nil {
// 		t.Fatal(err)
// 	}

// 	var xlsxBuf bytes.Buffer
// 	err = file.Write(&xlsxBuf)
// 	if err != nil {
// 		return
// 	}

// 	xlsxBytes := xlsxBuf.Bytes()

// 	targetDatas := make([]*baseExportModel, 0)
// 	err = ImportExcel(xlsxBytes, sheetName, &targetDatas)
// 	if err != nil {
// 		t.Fatal(err)
// 	}

// 	if !reflect.DeepEqual(originData, targetDatas) {
// 		t.Fatal(fmt.Errorf("origin target not equal \n origin:%+v \n target:%+v",
// 			originData, targetDatas))
// 	}
// }

// type timeExportModel struct {
// 	Time time.Time `excel:"index(0)"` // 精度会丢失
// }

// func TestTimeTypeExportExcel(t *testing.T) {
// 	sheetName := "sheet1"

// 	originData := make([]timeExportModel, 0)
// 	originData = append(originData, timeExportModel{
// 		Time: time.Now(),
// 	}, timeExportModel{
// 		Time: time.Now().Add(10000),
// 	})

// 	file := xlsx.NewFile()
// 	err := ExportExcel(file, sheetName, originData)
// 	if err != nil {
// 		t.Fatal(err)
// 	}

// 	var xlsxBuf bytes.Buffer
// 	err = file.Write(&xlsxBuf)
// 	if err != nil {
// 		return
// 	}

// 	xlsxBytes := xlsxBuf.Bytes()
// 	targetDatas := make([]timeExportModel, 0)
// 	err = ImportExcel(xlsxBytes, sheetName, &targetDatas)
// 	if err != nil {
// 		t.Fatal(err)
// 	}

// 	if len(originData) != len(targetDatas) {
// 		t.Fatal(fmt.Errorf("origin target len not equal"))
// 	}

// 	for i := 0; i < len(originData); i++ {
// 		oData := originData[i]
// 		tData := targetDatas[i]

// 		diffSeconds := oData.Time.Sub(tData.Time).Seconds()
// 		if math.Abs(diffSeconds) >= 1 {
// 			t.Fatal(fmt.Errorf("origin target time not equal"))
// 		}
// 	}

// }

// const letterBytes = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"

// func randStringBytes(n int) string {
// 	b := make([]byte, n)
// 	for i := range b {
// 		b[i] = letterBytes[rand.Intn(len(letterBytes))]
// 	}
// 	return string(b)
// }

// type sortExportModel struct {
// 	StringA string `excel:"index(4)"`
// 	StringB string `excel:"index(1)"`
// 	StringC string `excel:"index(0)"`
// }

// func TestSortExportExcel(t *testing.T) {
// 	sheetName := "sheet1"

// 	originData := make([]sortExportModel, 0)
// 	originData = append(originData, sortExportModel{
// 		StringA: randStringBytes(10),
// 		StringB: randStringBytes(10),
// 		StringC: randStringBytes(10),
// 	}, sortExportModel{
// 		StringA: randStringBytes(10),
// 		StringB: randStringBytes(10),
// 		StringC: randStringBytes(10),
// 	})

// 	file := xlsx.NewFile()
// 	err := ExportExcel(file, sheetName, &originData)
// 	if err != nil {
// 		t.Fatal(err)
// 	}

// 	var xlsxBuf bytes.Buffer
// 	err = file.Write(&xlsxBuf)
// 	if err != nil {
// 		return
// 	}

// 	xlsxBytes := xlsxBuf.Bytes()
// 	targetDatas := make([]sortExportModel, 0)
// 	err = ImportExcel(xlsxBytes, sheetName, &targetDatas)
// 	if err != nil {
// 		t.Fatal(err)
// 	}

// 	if !reflect.DeepEqual(originData, targetDatas) {
// 		t.Fatal(fmt.Errorf("origin target not equal \n origin:%+v \n target:%+v",
// 			originData, targetDatas))
// 	}
// }

// type formatExportOriginModel struct {
// 	Percent float64 `excel:"index(0);format(percent)"`
// }

// type formatExportTargetModel struct {
// 	PercentString string `excel:"index(0);"`
// }

// func TestFormatExportExcel(t *testing.T) {
// 	sheetName := "sheet1"

// 	originData := make([]formatExportOriginModel, 0)
// 	originData = append(originData, formatExportOriginModel{
// 		Percent: rand.Float64(),
// 	}, formatExportOriginModel{
// 		Percent: rand.Float64(),
// 	})

// 	file := xlsx.NewFile()
// 	err := ExportExcel(file, sheetName, originData)
// 	if err != nil {
// 		t.Fatal(err)
// 	}

// 	var xlsxBuf bytes.Buffer
// 	err = file.Write(&xlsxBuf)
// 	if err != nil {
// 		return
// 	}

// 	xlsxBytes := xlsxBuf.Bytes()
// 	targetDatas := make([]formatExportTargetModel, 0)
// 	err = ImportExcel(xlsxBytes, sheetName, &targetDatas)
// 	if err != nil {
// 		t.Fatal(err)
// 	}

// 	if len(originData) != len(targetDatas) {
// 		t.Fatal(fmt.Errorf("origin target len not equal"))
// 	}

// 	for i := 0; i < len(originData); i++ {
// 		oData := originData[i]
// 		tData := targetDatas[i]

// 		targetPercent := fmt.Sprintf("%.2f%%", oData.Percent)
// 		if targetPercent != tData.PercentString {
// 			t.Fatal(fmt.Errorf("complete_percent not equal data:%s xlxs:%s",
// 				targetPercent, tData.PercentString))
// 		}
// 	}
// }

// type headerExportModel struct {
// 	Header3 string `excel:"index(2);header(Header3)"`
// 	Header1 string `excel:"index(0);header(Header1)"`
// }

// func TestHeaderExportExcel(t *testing.T) {
// 	sheetName := "sheet1"

// 	originData := make([]headerExportModel, 0)

// 	file := xlsx.NewFile()
// 	err := ExportExcel(file, sheetName, originData)
// 	if err != nil {
// 		t.Fatal(err)
// 	}

// 	headRow := file.Sheet[sheetName].Rows[0]
// 	if headRow.Cells[0].String() != "Header1" ||
// 		headRow.Cells[2].String() != "Header3" {

// 		t.Fatal(fmt.Errorf("target head not equal headRow:%+v", headRow))
// 	}
// }

// func BenchmarkBaseTypeExportExcel(b *testing.B) {
// 	sheetName := "sheet1"
// 	originData := make([]baseExportModel, 0)

// 	for i := 0; i < b.N; i++ {
// 		originData = append(originData, baseExportData)
// 	}

// 	b.ResetTimer()
// 	file := xlsx.NewFile()
// 	ExportExcel(file, sheetName, originData)
// 	b.ReportAllocs()
// }
