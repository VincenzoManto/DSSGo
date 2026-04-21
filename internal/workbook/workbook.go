package workbook

import (
	"fmt"
	"regexp"
	"sort"
	"strconv"
	"strings"
)

var a1Pattern = regexp.MustCompile(`^([A-Z]+)([0-9]+)$`)

type Coord struct {
	Row int
	Col int
}

func (c Coord) A1() string {
	return c.ColumnName() + strconv.Itoa(c.Row+1)
}

func (c Coord) ColumnName() string {
	col := c.Col + 1
	if col <= 0 {
		return "A"
	}
	parts := []byte{}
	for col > 0 {
		col--
		parts = append([]byte{byte('A' + (col % 26))}, parts...)
		col /= 26
	}
	return string(parts)
}

func ParseA1(ref string) (Coord, error) {
	match := a1Pattern.FindStringSubmatch(strings.ToUpper(strings.TrimSpace(ref)))
	if match == nil {
		return Coord{}, fmt.Errorf("invalid A1 reference %q", ref)
	}
	row, err := strconv.Atoi(match[2])
	if err != nil || row <= 0 {
		return Coord{}, fmt.Errorf("invalid row in %q", ref)
	}
	col := 0
	for _, ch := range match[1] {
		col = (col * 26) + int(ch-'A'+1)
	}
	return Coord{Row: row - 1, Col: col - 1}, nil
}

func ParseRange(ref string) (Coord, Coord, error) {
	parts := strings.Split(strings.TrimSpace(ref), ":")
	if len(parts) == 1 {
		coord, err := ParseA1(parts[0])
		return coord, coord, err
	}
	if len(parts) != 2 {
		return Coord{}, Coord{}, fmt.Errorf("invalid range %q", ref)
	}
	start, err := ParseA1(parts[0])
	if err != nil {
		return Coord{}, Coord{}, err
	}
	end, err := ParseA1(parts[1])
	if err != nil {
		return Coord{}, Coord{}, err
	}
	if start.Row > end.Row {
		start.Row, end.Row = end.Row, start.Row
	}
	if start.Col > end.Col {
		start.Col, end.Col = end.Col, start.Col
	}
	return start, end, nil
}

type Cell struct {
	Input       string
	Value       string
	Error       string
	CachedValue string
}

func (c *Cell) IsFormula() bool {
	if c == nil {
		return false
	}
	return strings.HasPrefix(strings.TrimSpace(c.Input), "=")
}

type Sheet struct {
	Name  string
	Cells map[Coord]*Cell
}

func NewSheet(name string) *Sheet {
	return &Sheet{Name: name, Cells: map[Coord]*Cell{}}
}

func (s *Sheet) SetCell(row int, col int, input string) {
	coord := Coord{Row: row, Col: col}
	if strings.TrimSpace(input) == "" {
		delete(s.Cells, coord)
		return
	}
	cell := s.Cells[coord]
	if cell == nil {
		cell = &Cell{}
		s.Cells[coord] = cell
	}
	cell.Input = input
	cell.CachedValue = ""
	if !cell.IsFormula() {
		cell.Value = input
		cell.Error = ""
	}
	if cell.IsFormula() {
		cell.Error = ""
	}
	if input == "" {
		delete(s.Cells, coord)
	}
}

func (s *Sheet) GetCell(row int, col int) *Cell {
	return s.Cells[Coord{Row: row, Col: col}]
}

func (s *Sheet) SortedCoords() []Coord {
	coords := make([]Coord, 0, len(s.Cells))
	for coord := range s.Cells {
		coords = append(coords, coord)
	}
	sort.Slice(coords, func(i int, j int) bool {
		if coords[i].Row == coords[j].Row {
			return coords[i].Col < coords[j].Col
		}
		return coords[i].Row < coords[j].Row
	})
	return coords
}

// InsertRow shifts all cells at row >= row down by one.
func (s *Sheet) InsertRow(row int) {
	newCells := make(map[Coord]*Cell, len(s.Cells))
	for coord, cell := range s.Cells {
		if coord.Row >= row {
			newCells[Coord{Row: coord.Row + 1, Col: coord.Col}] = cell
		} else {
			newCells[coord] = cell
		}
	}
	s.Cells = newCells
}

// DeleteRow removes all cells at row and shifts cells below it up by one.
func (s *Sheet) DeleteRow(row int) {
	newCells := make(map[Coord]*Cell, len(s.Cells))
	for coord, cell := range s.Cells {
		if coord.Row < row {
			newCells[coord] = cell
		} else if coord.Row > row {
			newCells[Coord{Row: coord.Row - 1, Col: coord.Col}] = cell
		}
	}
	s.Cells = newCells
}

// InsertCol shifts all cells at col >= col right by one.
func (s *Sheet) InsertCol(col int) {
	newCells := make(map[Coord]*Cell, len(s.Cells))
	for coord, cell := range s.Cells {
		if coord.Col >= col {
			newCells[Coord{Row: coord.Row, Col: coord.Col + 1}] = cell
		} else {
			newCells[coord] = cell
		}
	}
	s.Cells = newCells
}

// DeleteCol removes all cells at col and shifts cells to the right of it left by one.
func (s *Sheet) DeleteCol(col int) {
	newCells := make(map[Coord]*Cell, len(s.Cells))
	for coord, cell := range s.Cells {
		if coord.Col < col {
			newCells[coord] = cell
		} else if coord.Col > col {
			newCells[Coord{Row: coord.Row, Col: coord.Col - 1}] = cell
		}
	}
	s.Cells = newCells
}

func (s *Sheet) Bounds() (minRow int, minCol int, maxRow int, maxCol int, ok bool) {
	if len(s.Cells) == 0 {
		return 0, 0, 0, 0, false
	}
	coords := s.SortedCoords()
	minRow, minCol = coords[0].Row, coords[0].Col
	maxRow, maxCol = coords[0].Row, coords[0].Col
	for _, coord := range coords[1:] {
		if coord.Row < minRow {
			minRow = coord.Row
		}
		if coord.Col < minCol {
			minCol = coord.Col
		}
		if coord.Row > maxRow {
			maxRow = coord.Row
		}
		if coord.Col > maxCol {
			maxCol = coord.Col
		}
	}
	return minRow, minCol, maxRow, maxCol, true
}

type Workbook struct {
	Metadata map[string]string
	Sheets   []*Sheet
}

func New() *Workbook {
	return &Workbook{Metadata: map[string]string{}, Sheets: []*Sheet{}}
}

func (w *Workbook) AddSheet(name string) *Sheet {
	sheet := NewSheet(name)
	w.Sheets = append(w.Sheets, sheet)
	return sheet
}

func (w *Workbook) SheetByName(name string) *Sheet {
	for _, sheet := range w.Sheets {
		if sheet.Name == name {
			return sheet
		}
	}
	return nil
}

func (w *Workbook) EnsureSheet(name string) *Sheet {
	if sheet := w.SheetByName(name); sheet != nil {
		return sheet
	}
	return w.AddSheet(name)
}
