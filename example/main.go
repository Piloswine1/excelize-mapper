package main

import (
	"time"

	excelizemapper "github.com/a631807682/excelize-mapper"

	"github.com/xuri/excelize/v2"
)

func main() {
	example1()
	example2()
}

type Sex int32

const (
	SexMale Sex = iota
	SexFemale
)

type User struct {
	ID        int
	Name      string    `excelize-mapper:"header:Name"`
	Desc      string    `excelize-mapper:"header:Desc;width:50"`
	Sex       Sex       `excelize-mapper:"header:Sex;format:sex"`
	Address   string    `excelize-mapper:"header:Address;default:China"`
	CreatedAt time.Time `excelize-mapper:"header:CreatedAt"`
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

func example1() {
	ct, _ := time.Parse(time.RFC3339, dt)
	data := []User{
		{
			ID:        1,
			Name:      "Tom",
			Desc:      "This is a long text, it will be wrapped.",
			Sex:       SexMale,
			Address:   "Singapore",
			CreatedAt: ct,
		}, {
			ID:        2,
			Name:      "Jerry",
			Desc:      "This is a long text.",
			Sex:       SexFemale,
			Address:   "",
			CreatedAt: ct,
		},
	}

	f := excelize.NewFile()
	defer f.Close()

	m := excelizemapper.NewExcelizeMapper(
		excelizemapper.WithFormatter("sex", SexFormat),
	)
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
	ID   int
	Name string `excelize-mapper:"header:Name;index:0;"`
	Desc string `excelize-mapper:"header:Desc;width:50;index:2"`
	Sex  Sex    `excelize-mapper:"header:Sex;format:sex;index:1"`
}

// custom index
func example2() {
	data := []User2{{
		ID:   1,
		Name: "Tom",
		Desc: "This is a long text, it will be wrapped.",
		Sex:  SexMale,
	}, {
		ID:   2,
		Name: "Jerry",
		Desc: "This is a long text.",
		Sex:  SexFemale,
	}}

	f := excelize.NewFile()
	defer f.Close()

	m := excelizemapper.NewExcelizeMapper(
		excelizemapper.WithAutoSort(false),
		excelizemapper.WithFormatter("sex", SexFormat),
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
