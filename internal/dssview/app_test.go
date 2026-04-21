package dssview

import (
	"testing"

	"github.com/vincm/dss/dss-go/internal/formula"
	"github.com/vincm/dss/dss-go/internal/workbook"
)

func TestMergeCachedFormulaValuesPreservesUnsupportedFormulaFallback(t *testing.T) {
	previous := workbook.New()
	prevSheet := previous.AddSheet("Sheet1")
	prevSheet.SetCell(0, 0, "=UNSUPPORTED(1)")
	prevCell := prevSheet.GetCell(0, 0)
	prevCell.CachedValue = "42"
	formula.RecalculateWorkbook(previous)

	next := workbook.New()
	nextSheet := next.AddSheet("Sheet1")
	nextSheet.SetCell(0, 0, "=UNSUPPORTED(1)")

	mergeCachedFormulaValues(previous, next)
	formula.RecalculateWorkbook(next)

	cell := nextSheet.GetCell(0, 0)
	if cell == nil {
		t.Fatal("expected formula cell to exist")
	}
	if cell.Value != "42" {
		t.Fatalf("expected cached fallback value 42, got %q", cell.Value)
	}
	if cell.Error != "" {
		t.Fatalf("expected no error after cached fallback, got %q", cell.Error)
	}
}
