package main

import (
	"errors"
	"fmt"

	"github.com/unidoc/unioffice/common"
	"github.com/unidoc/unioffice/measurement"
	"github.com/unidoc/unioffice/spreadsheet"
)

func main() {
	xlsxPresenter, err := NewXlsxPresenter()
	if err != nil {
		fmt.Println(err)
		return
	}

	workbook, err := xlsxPresenter.Run()
	if err != nil {
		fmt.Println(err)
		return
	}

	err = workbook.SaveToFile("/Users/kadyrbeknarmamatov/go/src/github.com/SEFI2/unioffice-run/document_output.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
}

type XlsxPresenter interface {
	Run() (*spreadsheet.Workbook, error)
}

type xlsxPresenterImpl struct {
	ss *spreadsheet.Workbook
}

func NewXlsxPresenter() (XlsxPresenter, error) {
	//ADD YOUR UNIOFFICE LICENSE HERE
	return &xlsxPresenterImpl{}, nil
}

func (xlsxPresenter *xlsxPresenterImpl) Run() (*spreadsheet.Workbook, error){
	xlsxPresenter.ss = spreadsheet.New()

	defer xlsxPresenter.ss.Close()

	xlsxPresenter.addLegend()

	if err := xlsxPresenter.ss.Validate(); err != nil {
		message := "Xlsx Presenter failed to validate the xlsx workbook"
		return nil, errors.New(message)
	}

	return xlsxPresenter.ss, nil
}

func (xlsxPresenter *xlsxPresenterImpl) addLegend() {
	legend := xlsxPresenter.ss.AddSheet()
	legend.SetName("Legend")

	xlsxPresenter.addLogo(&legend)
}

func (xlsxPresenter *xlsxPresenterImpl) addLogo(sheet *spreadsheet.Sheet) {
	logo, err := common.ImageFromFile("/Users/kadyrbeknarmamatov/go/src/github.com/SEFI2/unioffice-run/logo.png")
	if err != nil {
		return
	}

	iref, err := xlsxPresenter.ss.AddImage(logo)
	if err != nil {
		return
	} else {
		drawing := xlsxPresenter.ss.AddDrawing()
		sheet.SetDrawing(drawing)
		img := drawing.AddImage(iref, spreadsheet.AnchorTypeAbsolute)
		img.SetRowOffset(0)
		img.SetColOffset(0)
		var w measurement.Distance = 2.15 * measurement.Inch
		img.SetWidth(w)
		img.SetHeight(iref.RelativeHeight(w))
	}
}