package dss

import (
	"bytes"
	"encoding/csv"
	"fmt"
	"os"
	"regexp"
	"sort"
	"strings"

	"github.com/vincm/dss/dss-go/internal/workbook"
)

var bareLiteralPattern = regexp.MustCompile(`^\s*(-?\d+(\.\d+)?|(true|false|null))\s*$`)

func Parse(text string) (*workbook.Workbook, error) {
	text = strings.ReplaceAll(text, "\r\n", "\n")
	lines := strings.SplitAfter(text, "\n")
	wb := workbook.New()
	index := 0

	if len(lines) > 0 && strings.TrimSpace(lines[0]) == "---" {
		index++
		for index < len(lines) && strings.TrimSpace(lines[index]) != "---" {
			line := strings.TrimSpace(lines[index])
			if line != "" && strings.Contains(line, ":") {
				parts := strings.SplitN(line, ":", 2)
				wb.Metadata[strings.TrimSpace(parts[0])] = strings.TrimSpace(parts[1])
			}
			index++
		}
		if index < len(lines) && strings.TrimSpace(lines[index]) == "---" {
			index++
		}
	}

	var currentSheet *workbook.Sheet
	for index < len(lines) {
		line := strings.TrimRight(lines[index], "\n")
		trimmed := strings.TrimSpace(line)
		if trimmed == "" || strings.HasPrefix(trimmed, "#") {
			index++
			continue
		}
		if strings.HasPrefix(trimmed, "[") && strings.HasSuffix(trimmed, "]") {
			currentSheet = wb.EnsureSheet(strings.TrimSpace(trimmed[1 : len(trimmed)-1]))
			index++
			continue
		}
		if strings.HasPrefix(trimmed, "@") {
			if currentSheet == nil {
				return nil, fmt.Errorf("anchor declared before sheet at line %d", index+1)
			}
			anchorCoord, err := workbook.ParseA1(strings.TrimSpace(trimmed[1:]))
			if err != nil {
				return nil, err
			}
			block, nextIndex := collectAnchorBlock(lines, index+1)
			records, err := parseCSVBlock(block)
			if err != nil {
				return nil, err
			}
			for rowOffset, record := range records {
				for colOffset, value := range record {
					if strings.TrimSpace(value) == "" {
						continue
					}
					currentSheet.SetCell(anchorCoord.Row+rowOffset, anchorCoord.Col+colOffset, value)
				}
			}
			index = nextIndex
			continue
		}
		return nil, fmt.Errorf("unexpected content at line %d: %s", index+1, trimmed)
	}

	return wb, nil
}

func ReadFile(path string) (*workbook.Workbook, error) {
	content, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}
	return Parse(string(content))
}

func Serialize(wb *workbook.Workbook) (string, error) {
	var buf strings.Builder
	if len(wb.Metadata) > 0 {
		buf.WriteString("---\n")
		keys := make([]string, 0, len(wb.Metadata))
		for key := range wb.Metadata {
			keys = append(keys, key)
		}
		sort.Strings(keys)
		for _, key := range keys {
			buf.WriteString(key)
			buf.WriteString(": ")
			buf.WriteString(wb.Metadata[key])
			buf.WriteString("\n")
		}
		buf.WriteString("---\n")
	}

	for sheetIndex, sheet := range wb.Sheets {
		if sheetIndex > 0 {
			buf.WriteString("\n")
		}
		buf.WriteString("[")
		buf.WriteString(sheet.Name)
		buf.WriteString("]\n")
		blocks := buildAnchorBlocks(sheet)
		if len(blocks) == 0 {
			continue
		}
		for blockIndex, block := range blocks {
			if blockIndex > 0 {
				buf.WriteString("\n")
			}
			buf.WriteString("@ ")
			buf.WriteString(workbook.Coord{Row: block.startRow, Col: block.startCol}.A1())
			buf.WriteString("\n")
			for row := block.startRow; row <= block.endRow; row++ {
				cells := make([]string, block.endCol-block.startCol+1)
				for col := block.startCol; col <= block.endCol; col++ {
					cell := sheet.GetCell(row, col)
					if cell == nil {
						cells[col-block.startCol] = ""
						continue
					}
					cells[col-block.startCol] = encodeCell(cell.Input)
				}
				buf.WriteString(strings.Join(cells, ","))
				buf.WriteString("\n")
			}
		}
	}
	return buf.String(), nil
}

func WriteFile(path string, wb *workbook.Workbook) error {
	content, err := Serialize(wb)
	if err != nil {
		return err
	}
	return os.WriteFile(path, []byte(content), 0o644)
}


func WriteRaw(path string, text string) error {
	return os.WriteFile(path, []byte(text), 0o644)
}

func collectAnchorBlock(lines []string, start int) (string, int) {
	var block strings.Builder
	inQuotes := false
	for index := start; index < len(lines); index++ {
		line := strings.TrimRight(lines[index], "\n")
		trimmed := strings.TrimSpace(line)
		if !inQuotes && (strings.HasPrefix(trimmed, "[") && strings.HasSuffix(trimmed, "]") || strings.HasPrefix(trimmed, "@")) {
			return block.String(), index
		}
		if !inQuotes && (trimmed == "" || strings.HasPrefix(trimmed, "#")) {
			continue
		}
		block.WriteString(line)
		block.WriteString("\n")
		inQuotes = updateQuoteState(line, inQuotes)
	}
	return block.String(), len(lines)
}

func updateQuoteState(line string, current bool) bool {
	for index := 0; index < len(line); index++ {
		if line[index] != '"' {
			continue
		}
		if current && index+1 < len(line) && line[index+1] == '"' {
			index++
			continue
		}
		current = !current
	}
	return current
}

func parseCSVBlock(block string) ([][]string, error) {
	reader := csv.NewReader(strings.NewReader(block))
	reader.FieldsPerRecord = -1
	reader.TrimLeadingSpace = true
	reader.LazyQuotes = true
	return reader.ReadAll()
}

func encodeCell(value string) string {
	trimmed := strings.TrimSpace(value)
	if trimmed == "" {
		return ""
	}
	if strings.HasPrefix(trimmed, "=") {
		return quoteCSV(value)
	}
	if bareLiteralPattern.MatchString(trimmed) && !strings.ContainsAny(value, ",\n\"") {
		return trimmed
	}
	return quoteCSV(value)
}

func quoteCSV(value string) string {
	var buf bytes.Buffer
	writer := csv.NewWriter(&buf)
	if err := writer.Write([]string{value}); err != nil {
		return value
	}
	writer.Flush()
	return strings.TrimRight(buf.String(), "\n")
}

type anchorBlock struct {
	startRow int
	endRow   int
	startCol int
	endCol   int
}

type rowSegment struct {
	row      int
	startCol int
	endCol   int
}

func buildAnchorBlocks(sheet *workbook.Sheet) []anchorBlock {
	segments := buildRowSegments(sheet)
	if len(segments) == 0 {
		return nil
	}

	rows := make([]int, 0, len(segments))
	for row := range segments {
		rows = append(rows, row)
	}
	sort.Ints(rows)

	active := map[string]anchorBlock{}
	blocks := []anchorBlock{}
	for _, row := range rows {
		nextActive := map[string]anchorBlock{}
		for _, segment := range segments[row] {
			key := fmt.Sprintf("%d:%d", segment.startCol, segment.endCol)
			if current, ok := active[key]; ok && current.endRow == row-1 {
				current.endRow = row
				nextActive[key] = current
				continue
			}
			nextActive[key] = anchorBlock{
				startRow: segment.row,
				endRow:   segment.row,
				startCol: segment.startCol,
				endCol:   segment.endCol,
			}
		}
		for key, block := range active {
			if _, ok := nextActive[key]; !ok {
				blocks = append(blocks, block)
			}
		}
		active = nextActive
	}
	for _, block := range active {
		blocks = append(blocks, block)
	}

	sort.Slice(blocks, func(i int, j int) bool {
		if blocks[i].startRow == blocks[j].startRow {
			return blocks[i].startCol < blocks[j].startCol
		}
		return blocks[i].startRow < blocks[j].startRow
	})
	return blocks
}

func buildRowSegments(sheet *workbook.Sheet) map[int][]rowSegment {
	coords := sheet.SortedCoords()
	segments := map[int][]rowSegment{}
	if len(coords) == 0 {
		return segments
	}

	currentRow := coords[0].Row
	segmentStart := coords[0].Col
	segmentEnd := coords[0].Col
	for index := 1; index < len(coords); index++ {
		coord := coords[index]
		if coord.Row != currentRow || coord.Col > segmentEnd+1 {
			segments[currentRow] = append(segments[currentRow], rowSegment{row: currentRow, startCol: segmentStart, endCol: segmentEnd})
			currentRow = coord.Row
			segmentStart = coord.Col
			segmentEnd = coord.Col
			continue
		}
		segmentEnd = coord.Col
	}
	segments[currentRow] = append(segments[currentRow], rowSegment{row: currentRow, startCol: segmentStart, endCol: segmentEnd})
	return segments
}
