package formula

import (
	"testing"

	"github.com/vincm/dss/dss-go/internal/workbook"
)

func TestRecalculateWorkbook(t *testing.T) {
	wb := workbook.New()
	sheet := wb.AddSheet("Sheet1")
	sheet.SetCell(0, 0, "10")
	sheet.SetCell(1, 0, "20")
	sheet.SetCell(2, 0, "=SUM(A1:A2)")
	sheet.SetCell(3, 0, "=A3/2")

	RecalculateWorkbook(wb)

	if got := sheet.GetCell(2, 0).Value; got != "30" {
		t.Fatalf("A3 value mismatch: got %q", got)
	}
	if got := sheet.GetCell(3, 0).Value; got != "15" {
		t.Fatalf("A4 value mismatch: got %q", got)
	}
}

func TestIfAndComparisonSupport(t *testing.T) {
	wb := workbook.New()
	sheet := wb.AddSheet("Sheet1")
	sheet.SetCell(0, 0, "10")
	sheet.SetCell(0, 1, "=IF(A1>5,\"high\",\"low\")")
	sheet.SetCell(0, 2, "=CONCAT(B1,\"-\",ROUND(A1/3,1))")

	RecalculateWorkbook(wb)

	if got := sheet.GetCell(0, 1).Value; got != "high" {
		t.Fatalf("B1 value mismatch: got %q", got)
	}
	if got := sheet.GetCell(0, 2).Value; got != "high-3.3" {
		t.Fatalf("C1 value mismatch: got %q", got)
	}
}

func TestUnsupportedFunctionKeepsError(t *testing.T) {
	wb := workbook.New()
	sheet := wb.AddSheet("Sheet1")
	sheet.SetCell(0, 0, "=NOW()")

	RecalculateWorkbook(wb)

	if got := sheet.GetCell(0, 0).Value; got != "#ERR" {
		t.Fatalf("expected #ERR, got %q", got)
	}
	if sheet.GetCell(0, 0).Error == "" {
		t.Fatal("expected an evaluation error")
	}
}

func TestUnsupportedFunctionFallsBackToImportedCache(t *testing.T) {
	wb := workbook.New()
	sheet := wb.AddSheet("Sheet1")
	sheet.SetCell(0, 0, "=NOW()")
	sheet.GetCell(0, 0).CachedValue = "2026-04-21T12:00:00"

	RecalculateWorkbook(wb)

	if got := sheet.GetCell(0, 0).Value; got != "2026-04-21T12:00:00" {
		t.Fatalf("expected cached fallback value, got %q", got)
	}
	if got := sheet.GetCell(0, 0).Error; got != "" {
		t.Fatalf("expected no error when cached value is available, got %q", got)
	}
}
