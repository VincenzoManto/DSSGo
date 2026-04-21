package importers

import (
	"encoding/csv"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"github.com/extrame/xls"
	"github.com/xuri/excelize/v2"

	"github.com/vincm/dss/dss-go/internal/dss"
	"github.com/vincm/dss/dss-go/internal/workbook"
)

func LoadWorkbook(path string) (*workbook.Workbook, error) {
	switch strings.ToLower(filepath.Ext(path)) {
	case ".dss":
		return dss.ReadFile(path)
	case ".csv":
		return loadCSV(path)
	case ".xlsx", ".xlsm":
		return loadExcelize(path)
	case ".xls":
		return loadXLS(path)
	default:
		return nil, fmt.Errorf("unsupported file type %q", filepath.Ext(path))
	}
}

func loadCSV(path string) (*workbook.Workbook, error) {
	file, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	reader := csv.NewReader(file)
	reader.FieldsPerRecord = -1
	reader.TrimLeadingSpace = true

	wb := workbook.New()
	sheet := wb.AddSheet("Sheet1")
	rowIndex := 0
	for {
		record, err := reader.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		for colIndex, value := range record {
			if strings.TrimSpace(value) == "" {
				continue
			}
			sheet.SetCell(rowIndex, colIndex, value)
		}
		rowIndex++
	}
	return wb, nil
}

func loadExcelize(path string) (*workbook.Workbook, error) {
	file, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	wb := workbook.New()
	for _, name := range file.GetSheetList() {
		sheet := wb.AddSheet(name)
		dimension, err := file.GetSheetDimension(name)
		if err != nil || dimension == "" {
			continue
		}
		start, end, err := workbook.ParseRange(dimension)
		if err != nil {
			continue
		}
		for row := start.Row; row <= end.Row; row++ {
			for col := start.Col; col <= end.Col; col++ {
				ref := workbook.Coord{Row: row, Col: col}.A1()
				formulaText, _ := file.GetCellFormula(name, ref)
				value, _ := file.GetCellValue(name, ref)
				if formulaText == "" && value == "" {
					continue
				}
				input := value
				if formulaText != "" {
					input = "=" + formulaText
				}
				sheet.SetCell(row, col, input)
				if formulaText != "" {
					cell := sheet.GetCell(row, col)
					cell.Value = value
					cell.CachedValue = value
				}
			}
		}
	}
	return wb, nil
}

func loadXLS(path string) (*workbook.Workbook, error) {
	file, err := xls.Open(path, "utf-8")
	if err != nil {
		return nil, err
	}

	wb := workbook.New()
	for index := 0; index < file.NumSheets(); index++ {
		sheetSource := file.GetSheet(index)
		if sheetSource == nil {
			continue
		}
		sheetName := strings.TrimSpace(sheetSource.Name)
		if sheetName == "" {
			sheetName = fmt.Sprintf("Sheet%d", index+1)
		}
		sheet := wb.AddSheet(sheetName)
		for rowIndex := 0; rowIndex <= int(sheetSource.MaxRow); rowIndex++ {
			row := sheetSource.Row(rowIndex)
			if row == nil {
				continue
			}
			for colIndex := 0; colIndex < row.LastCol(); colIndex++ {
				value := row.Col(colIndex)
				if strings.TrimSpace(value) == "" {
					continue
				}
				sheet.SetCell(rowIndex, colIndex, value)
			}
		}
	}
	return wb, nil
}
