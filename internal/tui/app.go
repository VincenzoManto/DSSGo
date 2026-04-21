package tui

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"unicode"

	tea "github.com/charmbracelet/bubbletea"
	"github.com/charmbracelet/lipgloss"

	"github.com/vincm/dss/dss-go/internal/dashboard"
	"github.com/vincm/dss/dss-go/internal/dss"
	"github.com/vincm/dss/dss-go/internal/dssview"
	"github.com/vincm/dss/dss-go/internal/formula"
	"github.com/vincm/dss/dss-go/internal/importers"
	"github.com/vincm/dss/dss-go/internal/workbook"
)

const (
	rowLabelWidth = 6
	cellWidth     = 16
	minGridRows   = 5
)

// saveResultMsg is returned by the async save command.
type saveResultMsg struct{ err error }

// recalcResultMsg signals the async recalc is done.
type recalcResultMsg struct{}

func saveCmd(outputPath string, wb *workbook.Workbook) tea.Cmd {
	return func() tea.Msg {
		err := dss.WriteFile(outputPath, wb)
		return saveResultMsg{err: err}
	}
}

func recalcCmd(wb *workbook.Workbook) tea.Cmd {
	return func() tea.Msg {
		formula.RecalculateWorkbook(wb)
		return recalcResultMsg{}
	}
}

type model struct {
	workbook   *workbook.Workbook
	inputPath  string
	outputPath string
	sheetIndex int
	row        int
	col        int
	rowOffset  int
	colOffset  int
	width      int
	height     int
	modeIndex  int
	editMode   bool
	editBuffer string
	editCursor int
	// prompt overlay (rename sheet / new sheet)
	promptActive bool
	promptLabel  string
	promptBuffer string
	promptCursor int
	promptTarget string // "rename" or "new_sheet"
	status       string
	quit         bool
	openDash     bool
	openDSS      bool
}

var displayModes = []string{"values", "formulas", "both"}

var (
	titleStyle          = lipgloss.NewStyle().Bold(true).Foreground(lipgloss.Color("205"))
	selectedCellStyle   = lipgloss.NewStyle().Foreground(lipgloss.Color("229")).Background(lipgloss.Color("25")).Bold(true)
	formulaCellStyle    = lipgloss.NewStyle().Foreground(lipgloss.Color("86"))
	errorCellStyle      = lipgloss.NewStyle().Foreground(lipgloss.Color("203")).Bold(true)
	formulaBarCellStyle = lipgloss.NewStyle().Foreground(lipgloss.Color("117")).Bold(true)
)

func Run(inputPath string, outputPath string) error {
	wb, err := importers.LoadWorkbook(inputPath)
	if err != nil {
		return err
	}
	formula.RecalculateWorkbook(wb)
	if outputPath == "" {
		outputPath = defaultOutputPath(inputPath)
	}
	currentModel := model{
		workbook:   wb,
		inputPath:  inputPath,
		outputPath: outputPath,
		status:     "Arrows move, Enter edits, Tab changes sheet, d opens dashboard, Ctrl+S saves.",
	}
	for {
		program := tea.NewProgram(currentModel, tea.WithAltScreen())
		finalModel, err := program.Run()
		if err != nil {
			return err
		}
		m, ok := finalModel.(model)
		if !ok || !m.quit {
			return nil
		}

		if m.openDSS {
			res, err := dssview.Run(m.outputPath, m.workbook)
			if err != nil {
				return err
			}
			switch res.Action {
			case dssview.ExitToTUI:
				m.workbook = res.Workbook
				m.quit = false
				m.openDSS = false
				m.status = "Returned from DSS editor"
				currentModel = m
				continue
			case dssview.ExitToDashboard:
				action, err := dashboard.RunWorkbook(m.inputPath, res.Workbook)
				if err != nil {
					return err
				}
				if action != dashboard.ExitToTUI {
					return nil
				}
				m.workbook = res.Workbook
				m.quit = false
				m.openDSS = false
				m.openDash = false
				m.status = "Returned from dashboard"
				currentModel = m
				continue
			default:
				return nil
			}
		}

		if !m.openDash {
			return nil
		}
		action, err := dashboard.RunWorkbook(m.inputPath, m.workbook)
		if err != nil {
			return err
		}
		if action != dashboard.ExitToTUI {
			return nil
		}
		m.quit = false
		m.openDash = false
		m.status = "Returned from dashboard"
		currentModel = m
	}
}

func defaultOutputPath(inputPath string) string {
	ext := filepath.Ext(inputPath)
	if strings.EqualFold(ext, ".dss") {
		return inputPath
	}
	return strings.TrimSuffix(inputPath, ext) + ".dss"
}

func (m model) Init() tea.Cmd {
	return nil
}

func (m model) Update(msg tea.Msg) (tea.Model, tea.Cmd) {
	switch msg := msg.(type) {
	case tea.WindowSizeMsg:
		m.width = msg.Width
		m.height = msg.Height
		m.ensureVisible()
		return m, nil
	case saveResultMsg:
		if msg.err != nil {
			m.status = "Save failed: " + msg.err.Error()
		} else {
			m.status = "Saved to " + m.outputPath
		}
		return m, nil
	case recalcResultMsg:
		m.status = "Recalculated"
		return m, nil
	case tea.KeyMsg:
		if m.promptActive {
			return m.updatePrompt(msg)
		}
		if m.editMode {
			return m.updateEdit(msg)
		}
		return m.updateNormal(msg)
	}
	return m, nil
}

func (m model) updateNormal(msg tea.KeyMsg) (tea.Model, tea.Cmd) {
	switch msg.String() {
	case "ctrl+c", "q":
		m.quit = true
		return m, tea.Quit
	case "d":
		m.quit = true
		m.openDash = true
		return m, tea.Quit
	case "s":
		m.quit = true
		m.openDSS = true
		return m, tea.Quit
	case "up", "k":
		if m.row > 0 {
			m.row--
		}
	case "down", "j":
		m.row++
	case "left", "h":
		if m.col > 0 {
			m.col--
		}
	case "right", "l":
		m.col++
	case "enter", "e":
		m.beginEdit(currentCellInput(m), false, "Editing cell "+m.currentCoord().A1()+". Enter applies, Esc cancels.")
		return m, nil
	case "ctrl+s":
		m.status = "Saving…"
		m.ensureVisible()
		return m, saveCmd(m.outputPath, m.workbook)
	case "r":
		m.status = "Recalculating…"
		m.ensureVisible()
		return m, recalcCmd(m.workbook)
	case "tab", "]", "pgdown":
		if len(m.workbook.Sheets) > 0 {
			m.sheetIndex = (m.sheetIndex + 1) % len(m.workbook.Sheets)
			m.row, m.col, m.rowOffset, m.colOffset = 0, 0, 0, 0
			m.status = "Switched to sheet " + m.currentSheet().Name
		}
	case "shift+tab", "[", "pgup":
		if len(m.workbook.Sheets) > 0 {
			m.sheetIndex = (m.sheetIndex - 1 + len(m.workbook.Sheets)) % len(m.workbook.Sheets)
			m.row, m.col, m.rowOffset, m.colOffset = 0, 0, 0, 0
			m.status = "Switched to sheet " + m.currentSheet().Name
		}
	case "m":
		m.modeIndex = (m.modeIndex + 1) % len(displayModes)
		m.status = "Display mode: " + displayModes[m.modeIndex]
	case "x":
		coord := m.currentCoord().A1()
		m.currentSheet().SetCell(m.row, m.col, "")
		formula.RecalculateWorkbook(m.workbook)
		m.status = "Cleared " + coord
	case "o":
		m.currentSheet().InsertRow(m.row + 1)
		formula.RecalculateWorkbook(m.workbook)
		m.row++
		m.status = fmt.Sprintf("Inserted row %d", m.row+1)
	case "O":
		m.currentSheet().InsertRow(m.row)
		formula.RecalculateWorkbook(m.workbook)
		m.status = fmt.Sprintf("Inserted row %d above", m.row+1)
	case "D":
		m.currentSheet().DeleteRow(m.row)
		formula.RecalculateWorkbook(m.workbook)
		if m.row > 0 {
			m.row--
		}
		m.status = fmt.Sprintf("Deleted row %d", m.row+1)
	case "C":
		m.currentSheet().InsertCol(m.col)
		formula.RecalculateWorkbook(m.workbook)
		m.col++
		m.status = fmt.Sprintf("Inserted column %s", workbook.Coord{Col: m.col}.ColumnName())
	case "E":
		m.currentSheet().DeleteCol(m.col)
		formula.RecalculateWorkbook(m.workbook)
		if m.col > 0 {
			m.col--
		}
		m.status = fmt.Sprintf("Deleted column %s", workbook.Coord{Col: m.col}.ColumnName())
	case "N":
		m.promptActive = true
		m.promptLabel = "New sheet name"
		m.promptBuffer = ""
		m.promptCursor = 0
		m.promptTarget = "new_sheet"
		m.status = "Enter new sheet name, then press Enter"
	case "F2":
		m.promptActive = true
		m.promptLabel = "Rename sheet (was: " + m.currentSheet().Name + ")"
		m.promptBuffer = m.currentSheet().Name
		m.promptCursor = len([]rune(m.currentSheet().Name))
		m.promptTarget = "rename"
		m.status = "Enter new name for sheet, then press Enter"
	}
	if shouldStartDirectEdit(msg) {
		text := string(msg.Runes)
		m.beginEdit(text, true, "Editing cell "+m.currentCoord().A1()+". Typing replaced the current value.")
		return m, nil
	}
	if msg.Type == tea.KeySpace {
		m.beginEdit(" ", true, "Editing cell "+m.currentCoord().A1()+". Typing replaced the current value.")
		return m, nil
	}
	m.ensureVisible()
	return m, nil
}

func (m model) updateEdit(msg tea.KeyMsg) (tea.Model, tea.Cmd) {
	switch msg.String() {
	case "esc":
		m.editMode = false
		m.status = "Edit cancelled"
		return m, nil
	case "enter":
		m.currentSheet().SetCell(m.row, m.col, m.editBuffer)
		m.editMode = false
		m.status = "Updated " + m.currentCoord().A1()
		return m, recalcCmd(m.workbook)
	case "left", "ctrl+b":
		if m.editCursor > 0 {
			m.editCursor--
		}
		return m, nil
	case "right", "ctrl+f":
		if m.editCursor < len([]rune(m.editBuffer)) {
			m.editCursor++
		}
		return m, nil
	case "home", "ctrl+a":
		m.editCursor = 0
		return m, nil
	case "end", "ctrl+e":
		m.editCursor = len([]rune(m.editBuffer))
		return m, nil
	case "backspace", "ctrl+h":
		if m.editCursor > 0 {
			m.editBuffer = removeRuneAt(m.editBuffer, m.editCursor-1)
			m.editCursor--
		}
		return m, nil
	case "delete":
		if m.editCursor < len([]rune(m.editBuffer)) {
			m.editBuffer = removeRuneAt(m.editBuffer, m.editCursor)
		} else {
			m.editBuffer = ""
			m.editCursor = 0
		}
		return m, nil
	case "ctrl+u":
		m.editBuffer = ""
		m.editCursor = 0
		return m, nil
	case "up":
		step := m.editInnerWidth()
		if m.editCursor >= step {
			m.editCursor -= step
		} else {
			m.editCursor = 0
		}
		return m, nil
	case "down":
		step := m.editInnerWidth()
		next := m.editCursor + step
		max := len([]rune(m.editBuffer))
		if next > max {
			next = max
		}
		m.editCursor = next
		return m, nil
	}
	if msg.Type == tea.KeyRunes {
		m.editBuffer = insertAtCursor(m.editBuffer, m.editCursor, string(msg.Runes))
		m.editCursor += len(msg.Runes)
	}
	if msg.Type == tea.KeySpace {
		m.editBuffer = insertAtCursor(m.editBuffer, m.editCursor, " ")
		m.editCursor++
	}
	return m, nil
}

func (m model) View() string {
	if len(m.workbook.Sheets) == 0 {
		return "No sheets loaded\n"
	}
	var builder strings.Builder
	builder.WriteString(titleStyle.Render("DSS Terminal Spreadsheet"))
	builder.WriteString("\n")
	builder.WriteString(fitPlainLine(renderTabs(m.workbook.Sheets, m.sheetIndex), m.width))
	builder.WriteString("\n")
	builder.WriteString(fitPlainLine("File: "+m.inputPath+"  Save: "+m.outputPath, m.width))
	builder.WriteString(fitPlainLine(renderFormulaBar(m), m.width))
	builder.WriteString("\n")
	if m.editMode {
		builder.WriteString(fitPlainLine("Editing: type to update, Enter applies, Esc cancels", m.width))
		builder.WriteString("\n")
	} else {
		builder.WriteString(fitPlainLine(renderDashboardStrip(m), m.width))
		builder.WriteString("\n")
	}
	builder.WriteString(m.renderGrid())
	builder.WriteString("\n")
	if m.promptActive {
		builder.WriteString(renderPromptBox(m))
		builder.WriteString("\n")
	}
	builder.WriteString(fitPlainLine(m.status, m.width) + "\n")
	builder.WriteString(fitPlainLine("Keys: arrows move, type to edit, Enter edit current value, Tab/Shift+Tab sheet, d dashboard, s DSS source, Ctrl+S save, q quit", m.width) + "\n")
	builder.WriteString(fitPlainLine("Advanced: h/j/k/l move, r recalc, m mode, x clear, o/O rows, D delete row, C/E columns, N new sheet, F2 rename", m.width))
	return builder.String()
}

func renderTabs(sheets []*workbook.Sheet, current int) string {
	parts := make([]string, 0, len(sheets))
	for index, sheet := range sheets {
		if index == current {
			parts = append(parts, "["+sheet.Name+"]")
			continue
		}
		parts = append(parts, " "+sheet.Name+" ")
	}
	return strings.Join(parts, " | ")
}

func (m model) renderGrid() string {
	rowsVisible := m.visibleRowCount()
	colsVisible := m.visibleColCount()
	sheet := m.currentSheet()
	var builder strings.Builder
	builder.WriteString(renderHorizontalBorder(colsVisible))
	builder.WriteString("\n")
	builder.WriteString("│")
	builder.WriteString(padCenter("#", rowLabelWidth))
	builder.WriteString("│")
	for col := 0; col < colsVisible; col++ {
		label := workbook.Coord{Col: m.colOffset + col}.ColumnName()
		builder.WriteString(padCenter(label, cellWidth))
		builder.WriteString("│")
	}
	builder.WriteString("\n")
	builder.WriteString(renderHeaderDivider(colsVisible))
	builder.WriteString("\n")
	for row := 0; row < rowsVisible; row++ {
		absoluteRow := m.rowOffset + row
		builder.WriteString("│")
		builder.WriteString(padCenter(fmt.Sprintf("%d", absoluteRow+1), rowLabelWidth))
		builder.WriteString("│")
		for col := 0; col < colsVisible; col++ {
			absoluteCol := m.colOffset + col
			content := m.renderGridCell(absoluteRow, absoluteCol, sheet.GetCell(absoluteRow, absoluteCol))
			if absoluteRow == m.row && absoluteCol == m.col {
				content = styleSelectedGridContent(content, cellWidth)
			} else {
				content = styleGridContent(content, cellWidth)
			}
			builder.WriteString(content)
			builder.WriteString("│")
		}
		builder.WriteString("\n")
		if row == rowsVisible-1 {
			builder.WriteString(renderBottomBorder(colsVisible))
		} else {
			builder.WriteString(renderMiddleDivider(colsVisible))
		}
		builder.WriteString("\n")
	}
	return builder.String()
}

func (m model) renderGridCell(row int, col int, cell *workbook.Cell) string {
	if m.editMode && row == m.row && col == m.col {
		return m.editBuffer
	}
	if cell == nil {
		return ""
	}
	var content string
	switch displayModes[m.modeIndex] {
	case "values":
		if cell.Error != "" {
			content = cell.Value + " !"
		} else {
			content = cell.Value
		}
	case "formulas":
		content = cell.Input
	default:
		if cell.IsFormula() {
			content = cell.Input + " => " + cell.Value
		} else {
			content = cell.Value
		}
	}
	if cell.Error != "" {
		return errorCellStyle.Render(content)
	}
	if cell.IsFormula() {
		return formulaCellStyle.Render(content)
	}
	return content
}

func (m model) currentCellSummary() string {
	coord := m.currentCoord()
	cell := m.currentSheet().GetCell(m.row, m.col)
	if cell == nil {
		return fmt.Sprintf("Cell %s | input: <empty> | value: <empty>", coord.A1())
	}
	summary := fmt.Sprintf("Cell %s | input: %s | value: %s", coord.A1(), printable(cell.Input), printable(cell.Value))
	if cell.Error != "" {
		summary += " | error: " + cell.Error
	}
	return summary
}

func renderFormulaBar(m model) string {
	coord := m.currentCoord()
	cell := m.currentSheet().GetCell(m.row, m.col)
	available := m.width - 34
	if available < 18 {
		available = 18
	}
	inputWidth := available / 2
	valueWidth := available / 3
	errorWidth := available - inputWidth - valueWidth
	if errorWidth < 6 {
		errorWidth = 6
	}
	input := "<empty>"
	value := "<empty>"
	errorText := "none"
	if m.editMode {
		input = trimToWidth(printable(m.editBuffer), inputWidth)
		value = "editing..."
		return "Cell " + coord.A1() + " | Input: " + input + " | Value: " + trimToWidth(value, valueWidth) + " | Error: " + trimToWidth(errorText, errorWidth)
	}
	if cell != nil {
		input = printable(cell.Input)
		value = printable(cell.Value)
		if cell.Error != "" {
			errorText = cell.Error
		}
	}
	if cell != nil && cell.IsFormula() {
		input = formulaBarCellStyle.Render(input)
	}
	if cell != nil && cell.Error != "" {
		errorText = errorCellStyle.Render(errorText)
	}
	return "Cell " + coord.A1() + " | Input: " + trimToWidth(stripANSI(input), inputWidth) + " | Value: " + trimToWidth(stripANSI(value), valueWidth) + " | Error: " + trimToWidth(stripANSI(errorText), errorWidth)
}

func printable(value string) string {
	if value == "" {
		return "<empty>"
	}
	return value
}

func padRight(value string, width int) string {
	if runeLen(value) >= width {
		return value
	}
	return value + strings.Repeat(" ", width-runeLen(value))
}

func padCenter(value string, width int) string {
	length := runeLen(value)
	if length >= width {
		return trimToWidth(value, width)
	}
	left := (width - length) / 2
	right := width - length - left
	return strings.Repeat(" ", left) + value + strings.Repeat(" ", right)
}

func trimToWidth(value string, width int) string {
	if width <= 0 {
		return ""
	}
	runes := []rune(value)
	if len(runes) <= width {
		return value
	}
	if width == 1 {
		return string(runes[:1])
	}
	return string(runes[:width-1]) + "…"
}

func runeLen(value string) int {
	return len([]rune(value))
}

func insertAtCursor(value string, cursor int, addition string) string {
	runes := []rune(value)
	if cursor < 0 {
		cursor = 0
	}
	if cursor > len(runes) {
		cursor = len(runes)
	}
	prefix := string(runes[:cursor])
	suffix := string(runes[cursor:])
	return prefix + addition + suffix
}

func removeRuneAt(value string, index int) string {
	runes := []rune(value)
	if index < 0 || index >= len(runes) {
		return value
	}
	return string(runes[:index]) + string(runes[index+1:])
}

func renderEditBuffer(value string, cursor int) string {
	runes := []rune(value)
	if cursor < 0 {
		cursor = 0
	}
	if cursor > len(runes) {
		cursor = len(runes)
	}
	prefix := string(runes[:cursor])
	suffix := string(runes[cursor:])
	cursorGlyph := lipgloss.NewStyle().Reverse(true).Render(" ")
	if cursor < len(runes) {
		cursorGlyph = lipgloss.NewStyle().Reverse(true).Render(string(runes[cursor]))
		suffix = string(runes[cursor+1:])
	}
	if value == "" {
		return cursorGlyph
	}
	return prefix + cursorGlyph + suffix
}

func styleGridContent(value string, width int) string {
	plain := " " + trimToWidth(stripANSI(value), width-2)
	return padRight(plain, width)
}

func styleSelectedGridContent(value string, width int) string {
	plain := " " + trimToWidth(stripANSI(value), width-2)
	plain = padRight(plain, width)
	return selectedCellStyle.Render(plain)
}

func stripANSI(value string) string {
	builder := strings.Builder{}
	inEscape := false
	for _, r := range value {
		if r == '\u001b' {
			inEscape = true
			continue
		}
		if inEscape {
			if (r >= 'a' && r <= 'z') || (r >= 'A' && r <= 'Z') {
				inEscape = false
			}
			continue
		}
		builder.WriteRune(r)
	}
	return builder.String()
}

func renderHorizontalBorder(colsVisible int) string {
	var builder strings.Builder
	builder.WriteString("┌")
	builder.WriteString(strings.Repeat("─", rowLabelWidth))
	builder.WriteString("┬")
	for col := 0; col < colsVisible; col++ {
		builder.WriteString(strings.Repeat("─", cellWidth))
		if col == colsVisible-1 {
			builder.WriteString("┐")
		} else {
			builder.WriteString("┬")
		}
	}
	return builder.String()
}

func renderHeaderDivider(colsVisible int) string {
	var builder strings.Builder
	builder.WriteString("├")
	builder.WriteString(strings.Repeat("─", rowLabelWidth))
	builder.WriteString("┼")
	for col := 0; col < colsVisible; col++ {
		builder.WriteString(strings.Repeat("─", cellWidth))
		if col == colsVisible-1 {
			builder.WriteString("┤")
		} else {
			builder.WriteString("┼")
		}
	}
	return builder.String()
}

func renderMiddleDivider(colsVisible int) string {
	var builder strings.Builder
	builder.WriteString("├")
	builder.WriteString(strings.Repeat("─", rowLabelWidth))
	builder.WriteString("┼")
	for col := 0; col < colsVisible; col++ {
		builder.WriteString(strings.Repeat("─", cellWidth))
		if col == colsVisible-1 {
			builder.WriteString("┤")
		} else {
			builder.WriteString("┼")
		}
	}
	return builder.String()
}

func renderBottomBorder(colsVisible int) string {
	var builder strings.Builder
	builder.WriteString("└")
	builder.WriteString(strings.Repeat("─", rowLabelWidth))
	builder.WriteString("┴")
	for col := 0; col < colsVisible; col++ {
		builder.WriteString(strings.Repeat("─", cellWidth))
		if col == colsVisible-1 {
			builder.WriteString("┘")
		} else {
			builder.WriteString("┴")
		}
	}
	return builder.String()
}

type boxData struct {
	Title string
	Value string
}

func renderDashboardStrip(m model) string {
	filled := len(m.currentSheet().Cells)
	formulaCount := 0
	errorCount := 0
	for _, coord := range m.currentSheet().SortedCoords() {
		cell := m.currentSheet().GetCell(coord.Row, coord.Col)
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
	return fmt.Sprintf("Sheet: %s | Mode: %s | Cells: %d | Formulas: %d | Errors: %d", m.currentSheet().Name, displayModes[m.modeIndex], filled, formulaCount, errorCount)
}

func (m *model) ensureVisible() {
	rowsVisible := m.visibleRowCount()
	colsVisible := m.visibleColCount()
	if m.row < m.rowOffset {
		m.rowOffset = m.row
	}
	if m.row >= m.rowOffset+rowsVisible {
		m.rowOffset = m.row - rowsVisible + 1
	}
	if m.col < m.colOffset {
		m.colOffset = m.col
	}
	if m.col >= m.colOffset+colsVisible {
		m.colOffset = m.col - colsVisible + 1
	}
	if m.rowOffset < 0 {
		m.rowOffset = 0
	}
	if m.colOffset < 0 {
		m.colOffset = 0
	}
}

func (m model) visibleRowCount() int {
	return 10
}

func (m model) chromeHeight() int {
	height := 0
	height += 3 // title, tabs, file/save line
	height += 2 // formula bar + spacer
	if m.editMode {
		height += 2 // compact edit line + spacer
	} else {
		height += 2 // dashboard strip + spacer
	}
	if m.promptActive {
		height += 6 // prompt box + spacer
	}
	height += 3 // status + help lines
	return height
}

func (m model) visibleColCount() int {
	cols := (m.width - rowLabelWidth) / cellWidth
	if cols < 1 {
		cols = 1
	}
	return cols
}

func (m model) currentSheet() *workbook.Sheet {
	if len(m.workbook.Sheets) == 0 {
		return workbook.NewSheet("Sheet1")
	}
	if m.sheetIndex < 0 || m.sheetIndex >= len(m.workbook.Sheets) {
		return m.workbook.Sheets[0]
	}
	return m.workbook.Sheets[m.sheetIndex]
}

func (m model) currentCoord() workbook.Coord {
	return workbook.Coord{Row: m.row, Col: m.col}
}

func currentCellInput(m model) string {
	if cell := m.currentSheet().GetCell(m.row, m.col); cell != nil {
		return cell.Input
	}
	return ""
}

func shouldStartDirectEdit(msg tea.KeyMsg) bool {
	if msg.Type != tea.KeyRunes || len(msg.Runes) == 0 || msg.Alt {
		return false
	}
	for _, r := range msg.Runes {
		if !unicode.IsPrint(r) || unicode.IsControl(r) {
			return false
		}
	}
	return true
}

func fitPlainLine(value string, width int) string {
	if width <= 0 {
		return value
	}
	return trimToWidth(value, width)
}

func (m *model) beginEdit(initial string, replace bool, status string) {
	m.editMode = true
	if replace {
		m.editBuffer = initial
	} else {
		m.editBuffer = initial
	}
	m.editCursor = len([]rune(m.editBuffer))
	m.status = status
}
func (m model) editInnerWidth() int {
	w := m.width - 4
	if w < 20 {
		w = 20
	}
	return w
}

// renderPromptBox renders a centered prompt input overlay.
func renderPromptBox(m model) string {
	labelWidth := runeLen(m.promptLabel)
	bufDisplay := renderEditBuffer(m.promptBuffer, m.promptCursor)
	bufWidth := runeLen(stripANSI(m.promptBuffer))
	if bufWidth < 30 {
		bufWidth = 30
	}
	contentWidth := labelWidth
	if bufWidth > contentWidth {
		contentWidth = bufWidth
	}
	contentWidth += 4
	var builder strings.Builder
	builder.WriteString("┌")
	builder.WriteString(strings.Repeat("─", contentWidth))
	builder.WriteString("┐\n")
	builder.WriteString("│ ")
	builder.WriteString(padRight(m.promptLabel, contentWidth-2))
	builder.WriteString(" │\n")
	builder.WriteString("├")
	builder.WriteString(strings.Repeat("─", contentWidth))
	builder.WriteString("┤\n")
	builder.WriteString("│ ")
	builder.WriteString(padRight(bufDisplay, contentWidth-2))
	builder.WriteString(" │\n")
	builder.WriteString("└")
	builder.WriteString(strings.Repeat("─", contentWidth))
	builder.WriteString("┘")
	return builder.String()
}

func (m model) updatePrompt(msg tea.KeyMsg) (tea.Model, tea.Cmd) {
	switch msg.String() {
	case "esc":
		m.promptActive = false
		m.status = "Prompt cancelled"
		return m, nil
	case "enter":
		name := strings.TrimSpace(m.promptBuffer)
		if name == "" {
			m.status = "Name cannot be empty"
			m.promptActive = false
			return m, nil
		}
		switch m.promptTarget {
		case "new_sheet":
			m.workbook.AddSheet(name)
			m.sheetIndex = len(m.workbook.Sheets) - 1
			m.row, m.col, m.rowOffset, m.colOffset = 0, 0, 0, 0
			m.status = "Created sheet " + name
		case "rename":
			m.currentSheet().Name = name
			m.status = "Renamed sheet to " + name
		}
		m.promptActive = false
		return m, nil
	case "left", "ctrl+b":
		if m.promptCursor > 0 {
			m.promptCursor--
		}
	case "right", "ctrl+f":
		if m.promptCursor < len([]rune(m.promptBuffer)) {
			m.promptCursor++
		}
	case "home", "ctrl+a":
		m.promptCursor = 0
	case "end", "ctrl+e":
		m.promptCursor = len([]rune(m.promptBuffer))
	case "backspace", "ctrl+h":
		if m.promptCursor > 0 {
			m.promptBuffer = removeRuneAt(m.promptBuffer, m.promptCursor-1)
			m.promptCursor--
		}
	case "delete":
		if m.promptCursor < len([]rune(m.promptBuffer)) {
			m.promptBuffer = removeRuneAt(m.promptBuffer, m.promptCursor)
		}
	case "ctrl+u":
		m.promptBuffer = ""
		m.promptCursor = 0
	default:
		if msg.Type == tea.KeyRunes {
			m.promptBuffer = insertAtCursor(m.promptBuffer, m.promptCursor, string(msg.Runes))
			m.promptCursor += len(msg.Runes)
		}
		if msg.Type == tea.KeySpace {
			m.promptBuffer = insertAtCursor(m.promptBuffer, m.promptCursor, " ")
			m.promptCursor++
		}
	}
	return m, nil
}
func (m model) FileExists(path string) bool {
	_, err := os.Stat(path)
	return err == nil
}
