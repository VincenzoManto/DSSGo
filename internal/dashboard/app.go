package dashboard

import (
	"context"
	"fmt"
	"sort"
	"strings"

	"github.com/mum4k/termdash"
	"github.com/mum4k/termdash/container"
	"github.com/mum4k/termdash/keyboard"
	"github.com/mum4k/termdash/linestyle"
	"github.com/mum4k/termdash/terminal/tcell"
	"github.com/mum4k/termdash/terminal/terminalapi"
	"github.com/mum4k/termdash/widgets/barchart"
	textwidget "github.com/mum4k/termdash/widgets/text"

	"github.com/vincm/dss/dss-go/internal/formula"
	"github.com/vincm/dss/dss-go/internal/importers"
	"github.com/vincm/dss/dss-go/internal/workbook"
)

type ExitAction int

const (
	ExitClose ExitAction = iota
	ExitToTUI
)

func Run(inputPath string) (ExitAction, error) {
	wb, err := importers.LoadWorkbook(inputPath)
	if err != nil {
		return ExitClose, err
	}
	return RunWorkbook(inputPath, wb)
}

func RunWorkbook(inputPath string, wb *workbook.Workbook) (ExitAction, error) {
	formula.RecalculateWorkbook(wb)

	terminal, err := tcell.New()
	if err != nil {
		return ExitClose, err
	}
	defer terminal.Close()

	header, err := textwidget.New(textwidget.DisableScrolling(), textwidget.WrapAtWords())
	if err != nil {
		return ExitClose, err
	}
	if err := header.Write(buildHeader(inputPath, wb), textwidget.WriteReplace()); err != nil {
		return ExitClose, err
	}

	sheetsPanel, err := textwidget.New(textwidget.WrapAtWords())
	if err != nil {
		return ExitClose, err
	}
	if err := sheetsPanel.Write(buildSheetPanel(wb), textwidget.WriteReplace()); err != nil {
		return ExitClose, err
	}

	statsPanel, err := textwidget.New(textwidget.WrapAtWords())
	if err != nil {
		return ExitClose, err
	}
	if err := statsPanel.Write(buildStatsPanel(wb), textwidget.WriteReplace()); err != nil {
		return ExitClose, err
	}

	legendPanel, err := textwidget.New(textwidget.WrapAtWords())
	if err != nil {
		return ExitClose, err
	}
	if err := legendPanel.Write(buildLegendPanel(), textwidget.WriteReplace()); err != nil {
		return ExitClose, err
	}

	// Build barchart: filled cells per sheet
	cellValues, cellMax, cellLabels := buildCellsChart(wb)
	bc, err := barchart.New(
		barchart.ShowValues(),
		barchart.BarWidth(3),
		barchart.BarGap(1),
	)
	if err != nil {
		return ExitClose, err
	}
	if len(cellValues) > 0 {
		if err := bc.Values(cellValues, cellMax, barchart.Labels(cellLabels)); err != nil {
			return ExitClose, err
		}
	}
	root, err := container.New(
		terminal,
		container.Border(linestyle.Light),
		container.BorderTitle("DSS Dashboard"),
		container.SplitHorizontal(
			container.Top(
				container.Border(linestyle.Light),
				container.BorderTitle("Workbook"),
				container.PlaceWidget(header),
			),
			container.Bottom(
				container.SplitVertical(
					container.Left(
						container.Border(linestyle.Light),
						container.BorderTitle("Sheets"),
						container.PlaceWidget(sheetsPanel),
					),
					container.Right(
						container.SplitHorizontal(
							container.Top(
								container.SplitVertical(
									container.Left(
										container.Border(linestyle.Light),
										container.BorderTitle("Stats"),
										container.PlaceWidget(statsPanel),
									),
									container.Right(
										container.Border(linestyle.Light),
										container.BorderTitle("Cells / Sheet"),
										container.PlaceWidget(bc),
									),
									container.SplitPercent(55),
								),
							),
							container.Bottom(
								container.Border(linestyle.Light),
								container.BorderTitle("Controls"),
								container.PlaceWidget(legendPanel),
							),
							container.SplitPercent(70),
						),
					),
					container.SplitPercent(45),
				),
			),
			container.SplitFixed(5),
		),
	)
	if err != nil {
		return ExitClose, err
	}

	ctx, cancel := context.WithCancel(context.Background())
	defer cancel()
	action := ExitClose

	err = termdash.Run(ctx, terminal, root,
		termdash.KeyboardSubscriber(func(k *terminalapi.Keyboard) {
			if k.Key == keyboard.KeyEnter || k.Key == keyboard.Key('e') || k.Key == keyboard.Key('E') {
				action = ExitToTUI
				cancel()
				return
			}
			if k.Key == keyboard.KeyCtrlC || k.Key == keyboard.KeyCtrlQ || k.Key == keyboard.Key('q') || k.Key == keyboard.Key('Q') {
				cancel()
			}
		}),
		termdash.ErrorHandler(func(err error) {
			cancel()
		}),
	)
	return action, err
}

func buildHeader(inputPath string, wb *workbook.Workbook) string {
	formulaCount := 0
	errorCount := 0
	filledCells := 0
	for _, sheet := range wb.Sheets {
		for _, coord := range sheet.SortedCoords() {
			cell := sheet.GetCell(coord.Row, coord.Col)
			if cell == nil {
				continue
			}
			filledCells++
			if cell.IsFormula() {
				formulaCount++
			}
			if cell.Error != "" {
				errorCount++
			}
		}
	}
	return strings.Join([]string{
		"Path: " + inputPath,
		fmt.Sprintf("Sheets: %d", len(wb.Sheets)),
		fmt.Sprintf("Filled cells: %d", filledCells),
		fmt.Sprintf("Formulas: %d", formulaCount),
		fmt.Sprintf("Errors: %d", errorCount),
	}, "\n")
}

func buildSheetPanel(wb *workbook.Workbook) string {
	lines := make([]string, 0, len(wb.Sheets))
	for _, sheet := range wb.Sheets {
		minRow, minCol, maxRow, maxCol, ok := sheet.Bounds()
		if !ok {
			lines = append(lines, sheet.Name+": empty")
			continue
		}
		lines = append(lines, fmt.Sprintf(
			"%s\n  range: %s -> %s\n  cells: %d",
			sheet.Name,
			workbook.Coord{Row: minRow, Col: minCol}.A1(),
			workbook.Coord{Row: maxRow, Col: maxCol}.A1(),
			len(sheet.Cells),
		))
	}
	if len(lines) == 0 {
		return "No sheets"
	}
	return strings.Join(lines, "\n\n")
}

func buildStatsPanel(wb *workbook.Workbook) string {
	topSheets := make([]string, 0, len(wb.Sheets))
	for _, sheet := range wb.Sheets {
		formulaCount := 0
		errorCount := 0
		for _, coord := range sheet.SortedCoords() {
			cell := sheet.GetCell(coord.Row, coord.Col)
			if cell == nil {
				continue
			}
			if cell.IsFormula() {
				formulaCount++
			}
			if cell.Error != "" {
				errorCount++
			}
		}
		topSheets = append(topSheets, fmt.Sprintf("%s: %d formulas, %d errors", sheet.Name, formulaCount, errorCount))
	}
	sort.Strings(topSheets)
	metadataKeys := make([]string, 0, len(wb.Metadata))
	for key := range wb.Metadata {
		metadataKeys = append(metadataKeys, key)
	}
	sort.Strings(metadataKeys)
	metadataSummary := "none"
	if len(metadataKeys) > 0 {
		metadataSummary = strings.Join(metadataKeys, ", ")
	}
	return "Metadata keys: " + metadataSummary + "\n\n" + strings.Join(topSheets, "\n")
}

func buildLegendPanel() string {
	return strings.Join([]string{
		"Enter / e: open the editable spreadsheet view",
		"q / Ctrl+Q / Ctrl+C: close dashboard",
		"Use `dss <file>` for the editable spreadsheet TUI",
		"Use `dss convert input output.dss` to export directly",
	}, "\n")
}

// buildCellsChart returns values, max, and labels for a per-sheet cells barchart.
func buildCellsChart(wb *workbook.Workbook) (values []int, max int, labels []string) {
	max = 1
	for _, sheet := range wb.Sheets {
		count := len(sheet.Cells)
		values = append(values, count)
		labels = append(labels, sheet.Name)
		if count > max {
			max = count
		}
	}
	return values, max, labels
}
