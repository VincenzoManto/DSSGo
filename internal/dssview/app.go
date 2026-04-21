package dssview

import (
	"strings"
	"unicode"

	tea "github.com/charmbracelet/bubbletea"
	"github.com/charmbracelet/lipgloss"

	"github.com/vincm/dss/dss-go/internal/dss"
	"github.com/vincm/dss/dss-go/internal/formula"
	"github.com/vincm/dss/dss-go/internal/workbook"
)


type ExitAction int

const (
	ExitClose ExitAction = iota
	ExitToTUI
	ExitToDashboard
)


type Result struct {
	Action   ExitAction
	Workbook *workbook.Workbook 
}

type saveResultMsg struct{ err error }
type parseResultMsg struct {
	wb  *workbook.Workbook
	err error
}

func mergeCachedFormulaValues(previous *workbook.Workbook, next *workbook.Workbook) {
	if previous == nil || next == nil {
		return
	}
	for _, nextSheet := range next.Sheets {
		prevSheet := previous.SheetByName(nextSheet.Name)
		if prevSheet == nil {
			continue
		}
		for _, coord := range nextSheet.SortedCoords() {
			nextCell := nextSheet.GetCell(coord.Row, coord.Col)
			prevCell := prevSheet.GetCell(coord.Row, coord.Col)
			if nextCell == nil || prevCell == nil {
				continue
			}
			if !nextCell.IsFormula() || !prevCell.IsFormula() {
				continue
			}
			if strings.TrimSpace(nextCell.Input) != strings.TrimSpace(prevCell.Input) {
				continue
			}
			if prevCell.CachedValue != "" {
				nextCell.CachedValue = prevCell.CachedValue
				continue
			}
			if prevCell.Value != "" && prevCell.Error == "" {
				nextCell.CachedValue = prevCell.Value
			}
		}
	}
}



func saveAndParseCmd(outputPath string, previous *workbook.Workbook, lines []string) tea.Cmd {
	return func() tea.Msg {
		text := strings.Join(lines, "\n")
		if err := dss.WriteRaw(outputPath, text); err != nil {
			return saveResultMsg{err: err}
		}
		wb, err := dss.Parse(text)
		if err != nil {
			return parseResultMsg{err: err}
		}
		mergeCachedFormulaValues(previous, wb)
		formula.RecalculateWorkbook(wb)
		return parseResultMsg{wb: wb}
	}
}

type model struct {
	workbook   *workbook.Workbook
	outputPath string
	lines      []string
	lineIdx    int
	colIdx     int
	scrollTop  int
	width      int
	height     int
	dirty      bool
	status     string
	action     ExitAction
	done       bool
}

var (
	titleStyle    = lipgloss.NewStyle().Bold(true).Foreground(lipgloss.Color("205"))
	curLineStyle  = lipgloss.NewStyle().Background(lipgloss.Color("237"))
	dirtyStyle    = lipgloss.NewStyle().Foreground(lipgloss.Color("203")).Bold(true)
	modifiedStyle = lipgloss.NewStyle().Foreground(lipgloss.Color("214"))
)

func Run(outputPath string, wb *workbook.Workbook) (Result, error) {
	
	text, err := dss.Serialize(wb)
	if err != nil {
		return Result{Action: ExitClose, Workbook: wb}, err
	}
	lines := strings.Split(strings.ReplaceAll(text, "\r\n", "\n"), "\n")

	m := model{
		workbook:   wb,
		outputPath: outputPath,
		lines:      lines,
		status:     "Ctrl+S save & apply, e/Esc back to spreadsheet, d dashboard, q quit",
	}
	program := tea.NewProgram(m, tea.WithAltScreen())
	finalRaw, err := program.Run()
	if err != nil {
		return Result{Action: ExitClose, Workbook: wb}, err
	}
	final := finalRaw.(model)
	returnWB := final.workbook
	if returnWB == nil {
		returnWB = wb
	}
	return Result{Action: final.action, Workbook: returnWB}, nil
}

func (m model) Init() tea.Cmd { return nil }

func (m model) Update(msg tea.Msg) (tea.Model, tea.Cmd) {
	switch msg := msg.(type) {
	case tea.WindowSizeMsg:
		m.width = msg.Width
		m.height = msg.Height

	case saveResultMsg:
		if msg.err != nil {
			m.status = "Save error: " + msg.err.Error()
		} else {
			m.status = "Saved to " + m.outputPath
		}

	case parseResultMsg:
		if msg.err != nil {
			m.status = "Parse error: " + msg.err.Error()
		} else {
			m.workbook = msg.wb
			m.dirty = false
			m.status = "Saved and applied"
		}

	case tea.KeyMsg:
		return m.handleKey(msg)
	}
	return m, nil
}

func (m model) handleKey(msg tea.KeyMsg) (tea.Model, tea.Cmd) {
	switch msg.String() {
	case "ctrl+c", "q":
		m.done = true
		m.action = ExitClose
		return m, tea.Quit

	case "e", "esc":
		m.done = true
		m.action = ExitToTUI
		return m, tea.Quit

	case "d":
		m.done = true
		m.action = ExitToDashboard
		return m, tea.Quit

	case "ctrl+s":
		m.status = "Saving…"
		return m, saveAndParseCmd(m.outputPath, m.workbook, m.lines)

	case "up", "k":
		if m.lineIdx > 0 {
			m.lineIdx--
			m.clampCol()
			m.ensureVisible()
		}

	case "down", "j":
		if m.lineIdx < len(m.lines)-1 {
			m.lineIdx++
			m.clampCol()
			m.ensureVisible()
		}

	case "left":
		if m.colIdx > 0 {
			m.colIdx--
		} else if m.lineIdx > 0 {
			m.lineIdx--
			m.colIdx = len([]rune(m.lines[m.lineIdx]))
			m.ensureVisible()
		}

	case "right":
		lineRunes := []rune(m.lines[m.lineIdx])
		if m.colIdx < len(lineRunes) {
			m.colIdx++
		} else if m.lineIdx < len(m.lines)-1 {
			m.lineIdx++
			m.colIdx = 0
			m.ensureVisible()
		}

	case "home", "ctrl+a":
		m.colIdx = 0

	case "end", "ctrl+e":
		m.colIdx = len([]rune(m.lines[m.lineIdx]))

	case "pgdown":
		step := m.visibleLines()
		m.lineIdx += step
		if m.lineIdx >= len(m.lines) {
			m.lineIdx = len(m.lines) - 1
		}
		m.clampCol()
		m.ensureVisible()

	case "pgup":
		step := m.visibleLines()
		m.lineIdx -= step
		if m.lineIdx < 0 {
			m.lineIdx = 0
		}
		m.clampCol()
		m.ensureVisible()

	case "enter":
		
		lineRunes := []rune(m.lines[m.lineIdx])
		before := string(lineRunes[:m.colIdx])
		after := string(lineRunes[m.colIdx:])
		newLines := make([]string, 0, len(m.lines)+1)
		newLines = append(newLines, m.lines[:m.lineIdx]...)
		newLines = append(newLines, before, after)
		newLines = append(newLines, m.lines[m.lineIdx+1:]...)
		m.lines = newLines
		m.lineIdx++
		m.colIdx = 0
		m.dirty = true
		m.ensureVisible()

	case "backspace", "ctrl+h":
		lineRunes := []rune(m.lines[m.lineIdx])
		if m.colIdx > 0 {
			newLine := string(lineRunes[:m.colIdx-1]) + string(lineRunes[m.colIdx:])
			m.lines[m.lineIdx] = newLine
			m.colIdx--
			m.dirty = true
		} else if m.lineIdx > 0 {
			prevRunes := []rune(m.lines[m.lineIdx-1])
			m.colIdx = len(prevRunes)
			merged := string(prevRunes) + string(lineRunes)
			newLines := make([]string, 0, len(m.lines)-1)
			newLines = append(newLines, m.lines[:m.lineIdx-1]...)
			newLines = append(newLines, merged)
			newLines = append(newLines, m.lines[m.lineIdx+1:]...)
			m.lines = newLines
			m.lineIdx--
			m.dirty = true
			m.ensureVisible()
		}

	case "delete":
		lineRunes := []rune(m.lines[m.lineIdx])
		if m.colIdx < len(lineRunes) {
			newLine := string(lineRunes[:m.colIdx]) + string(lineRunes[m.colIdx+1:])
			m.lines[m.lineIdx] = newLine
			m.dirty = true
		} else if m.lineIdx < len(m.lines)-1 {
			merged := m.lines[m.lineIdx] + m.lines[m.lineIdx+1]
			newLines := make([]string, 0, len(m.lines)-1)
			newLines = append(newLines, m.lines[:m.lineIdx]...)
			newLines = append(newLines, merged)
			newLines = append(newLines, m.lines[m.lineIdx+2:]...)
			m.lines = newLines
			m.dirty = true
		}

	default:
		if msg.Type == tea.KeyRunes && !msg.Alt {
			for _, r := range msg.Runes {
				if !unicode.IsPrint(r) || unicode.IsControl(r) {
					return m, nil
				}
			}
			lineRunes := []rune(m.lines[m.lineIdx])
			ins := []rune(string(msg.Runes))
			newLine := string(lineRunes[:m.colIdx]) + string(ins) + string(lineRunes[m.colIdx:])
			m.lines[m.lineIdx] = newLine
			m.colIdx += len(ins)
			m.dirty = true
		}
		if msg.Type == tea.KeySpace {
			lineRunes := []rune(m.lines[m.lineIdx])
			newLine := string(lineRunes[:m.colIdx]) + " " + string(lineRunes[m.colIdx:])
			m.lines[m.lineIdx] = newLine
			m.colIdx++
			m.dirty = true
		}
	}
	return m, nil
}

func (m model) View() string {
	var b strings.Builder

	dirtyMark := ""
	if m.dirty {
		dirtyMark = dirtyStyle.Render(" [modified]")
	}
	b.WriteString(titleStyle.Render("DSS Source Editor"))
	b.WriteString(dirtyMark)
	b.WriteString("\n")
	b.WriteString(trimLine("File: "+m.outputPath, m.width))
	b.WriteString("\n")

	visLines := m.visibleLines()
	for i := 0; i < visLines; i++ {
		lineNum := m.scrollTop + i
		if lineNum >= len(m.lines) {
			b.WriteString("\n")
			continue
		}
		lineText := m.lines[lineNum]
		lineRunes := []rune(lineText)

		var rendered string
		if lineNum == m.lineIdx {
			
			col := m.colIdx
			if col > len(lineRunes) {
				col = len(lineRunes)
			}
			before := string(lineRunes[:col])
			cursorChar := " "
			after := ""
			if col < len(lineRunes) {
				cursorChar = string(lineRunes[col])
				after = string(lineRunes[col+1:])
			}
			cursor := lipgloss.NewStyle().Reverse(true).Render(cursorChar)
			raw := before + cursor + after
			rendered = curLineStyle.Render(trimLine(raw, m.width))
		} else {
			rendered = trimLine(lineText, m.width)
		}
		b.WriteString(rendered)
		b.WriteString("\n")
	}

	b.WriteString(trimLine(m.status, m.width))
	b.WriteString("\n")
	b.WriteString(trimLine(
		"Keys: arrows navigate, type to edit, Enter new line, Ctrl+S save & apply, e back to spreadsheet, d dashboard, q quit",
		m.width,
	))
	return b.String()
}

func (m model) visibleLines() int {
	lines := m.height - 5 
	if lines < 5 {
		lines = 5
	}
	return lines
}

func (m model) clampCol() {
	max := len([]rune(m.lines[m.lineIdx]))
	if m.colIdx > max {
		m.colIdx = max
	}
}

func (m *model) ensureVisible() {
	visLines := m.visibleLines()
	if m.lineIdx < m.scrollTop {
		m.scrollTop = m.lineIdx
	}
	if m.lineIdx >= m.scrollTop+visLines {
		m.scrollTop = m.lineIdx - visLines + 1
	}
}

func trimLine(s string, width int) string {
	if width <= 0 {
		return s
	}
	runes := []rune(s)
	if len(runes) <= width {
		return s
	}
	if width <= 1 {
		return string(runes[:width])
	}
	return string(runes[:width-1]) + "…"
}
