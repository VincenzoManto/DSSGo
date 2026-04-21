package dss

import (
	"strings"
	"testing"

	"github.com/vincm/dss/dss-go/internal/workbook"
)

func TestParseAndSerialize(t *testing.T) {
	input := strings.Join([]string{
		"---",
		"project: Demo",
		"---",
		"[Sheet1]",
		"@ B2",
		"1,2,\"=SUM(B2:C2)\"",
		"\"hello\",\"multiline",
		"value\"",
	}, "\n")

	wb, err := Parse(input)
	if err != nil {
		t.Fatalf("Parse() error = %v", err)
	}
	if got := wb.Metadata["project"]; got != "Demo" {
		t.Fatalf("metadata mismatch: got %q", got)
	}
	sheet := wb.SheetByName("Sheet1")
	if sheet == nil {
		t.Fatal("expected Sheet1 to exist")
	}
	if got := sheet.GetCell(1, 1).Input; got != "1" {
		t.Fatalf("B2 mismatch: got %q", got)
	}
	if got := sheet.GetCell(1, 3).Input; got != "=SUM(B2:C2)" {
		t.Fatalf("D2 mismatch: got %q", got)
	}
	if got := sheet.GetCell(2, 2).Input; got != "multiline\nvalue" {
		t.Fatalf("C3 mismatch: got %q", got)
	}

	serialized, err := Serialize(wb)
	if err != nil {
		t.Fatalf("Serialize() error = %v", err)
	}
	if !strings.Contains(serialized, "[Sheet1]") || !strings.Contains(serialized, "@ B2") {
		t.Fatalf("serialized output missing expected DSS structure: %s", serialized)
	}
}

func TestSerializeUsesBoundingAnchor(t *testing.T) {
	wb := workbook.New()
	sheet := wb.AddSheet("Sheet1")
	sheet.SetCell(3, 3, "42")

	serialized, err := Serialize(wb)
	if err != nil {
		t.Fatalf("Serialize() error = %v", err)
	}
	if !strings.Contains(serialized, "@ D4") {
		t.Fatalf("expected anchor at D4, got %s", serialized)
	}
}

func TestSerializeUsesSparseAnchors(t *testing.T) {
	wb := workbook.New()
	sheet := wb.AddSheet("Sheet1")
	sheet.SetCell(0, 0, "1")
	sheet.SetCell(0, 1, "2")
	sheet.SetCell(5, 4, "tail")

	serialized, err := Serialize(wb)
	if err != nil {
		t.Fatalf("Serialize() error = %v", err)
	}
	if !strings.Contains(serialized, "@ A1\n1,2\n\n@ E6\ntail\n") {
		t.Fatalf("expected sparse anchor layout, got %s", serialized)
	}
}
