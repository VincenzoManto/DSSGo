package formula

import (
	"fmt"
	"math"
	"strconv"
	"strings"
	"unicode"
	"unicode/utf8"

	"github.com/vincm/dss/dss-go/internal/workbook"
)

type tokenKind int

const (
	tokenEOF tokenKind = iota
	tokenNumber
	tokenString
	tokenIdentifier
	tokenCell
	tokenPlus
	tokenMinus
	tokenMul
	tokenDiv
	tokenLParen
	tokenRParen
	tokenComma
	tokenColon
	tokenEq
	tokenNE
	tokenLT
	tokenLTE
	tokenGT
	tokenGTE
)

type token struct {
	kind tokenKind
	text string
}

type valueKind int

const (
	valueBlank valueKind = iota
	valueNumber
	valueString
	valueBool
	valueRange
)

type formulaValue struct {
	kind   valueKind
	number float64
	text   string
	boolv  bool
	rangev []formulaValue
}

type evaluator struct {
	sheet    *workbook.Sheet
	cache    map[workbook.Coord]string
	visiting map[workbook.Coord]bool
}

func RecalculateWorkbook(wb *workbook.Workbook) {
	for _, sheet := range wb.Sheets {
		e := &evaluator{
			sheet:    sheet,
			cache:    map[workbook.Coord]string{},
			visiting: map[workbook.Coord]bool{},
		}
		for _, coord := range sheet.SortedCoords() {
			cell := sheet.GetCell(coord.Row, coord.Col)
			if cell == nil {
				continue
			}
			if !cell.IsFormula() {
				cell.Value = cell.Input
				cell.Error = ""
				continue
			}
			value, err := e.evalCell(coord)
			if err != nil {
				if cell.Value == "" {
					cell.Value = "#ERR"
				}
				cell.Error = err.Error()
				continue
			}
			cell.Value = value
			cell.Error = ""
		}
	}
}

func (e *evaluator) evalCell(coord workbook.Coord) (string, error) {
	if cached, ok := e.cache[coord]; ok {
		return cached, nil
	}
	if e.visiting[coord] {
		return "", fmt.Errorf("cyclic reference at %s", coord.A1())
	}
	cell := e.sheet.GetCell(coord.Row, coord.Col)
	if cell == nil {
		return "", nil
	}
	if !cell.IsFormula() {
		return cell.Input, nil
	}

	e.visiting[coord] = true
	defer delete(e.visiting, coord)

	parser := newParser(strings.TrimSpace(strings.TrimPrefix(cell.Input, "=")), e)
	value, err := parser.parse()
	if err != nil {
		if cell.CachedValue != "" {
			e.cache[coord] = cell.CachedValue
			return cell.CachedValue, nil
		}
		return "", err
	}
	result := value.outputString()
	e.cache[coord] = result
	return result, nil
}

func (e *evaluator) evalCellValue(coord workbook.Coord) (formulaValue, error) {
	value, err := e.evalCell(coord)
	if err != nil {
		return formulaValue{}, err
	}
	return parseLiteralValue(value), nil
}

func (e *evaluator) evalRange(start workbook.Coord, end workbook.Coord) ([]formulaValue, error) {
	values := []formulaValue{}
	for row := start.Row; row <= end.Row; row++ {
		for col := start.Col; col <= end.Col; col++ {
			value, err := e.evalCellValue(workbook.Coord{Row: row, Col: col})
			if err != nil {
				return nil, err
			}
			if value.kind == valueBlank {
				continue
			}
			values = append(values, value)
		}
	}
	return values, nil
}

type parser struct {
	tokens []token
	index  int
	eval   *evaluator
}

func newParser(input string, eval *evaluator) *parser {
	return &parser{tokens: tokenize(input), eval: eval}
}

func (p *parser) parse() (formulaValue, error) {
	value, err := p.parseComparison()
	if err != nil {
		return formulaValue{}, err
	}
	if p.peek().kind != tokenEOF {
		return formulaValue{}, fmt.Errorf("unexpected token %q", p.peek().text)
	}
	if value.kind == valueRange {
		return formulaValue{}, fmt.Errorf("range used where scalar value was expected")
	}
	return value, nil
}

func (p *parser) parseComparison() (formulaValue, error) {
	left, err := p.parseExpression()
	if err != nil {
		return formulaValue{}, err
	}
	for {
		switch p.peek().kind {
		case tokenEq, tokenNE, tokenLT, tokenLTE, tokenGT, tokenGTE:
			op := p.next()
			right, err := p.parseExpression()
			if err != nil {
				return formulaValue{}, err
			}
			result, err := compareValues(left, right, op.kind)
			if err != nil {
				return formulaValue{}, err
			}
			left = boolValue(result)
		default:
			return left, nil
		}
	}
}

func (p *parser) parseExpression() (formulaValue, error) {
	left, err := p.parseTerm()
	if err != nil {
		return formulaValue{}, err
	}
	for {
		switch p.peek().kind {
		case tokenPlus:
			p.next()
			right, err := p.parseTerm()
			if err != nil {
				return formulaValue{}, err
			}
			left, err = combineNumeric(left, right, func(a float64, b float64) float64 { return a + b })
			if err != nil {
				return formulaValue{}, err
			}
		case tokenMinus:
			p.next()
			right, err := p.parseTerm()
			if err != nil {
				return formulaValue{}, err
			}
			left, err = combineNumeric(left, right, func(a float64, b float64) float64 { return a - b })
			if err != nil {
				return formulaValue{}, err
			}
		default:
			return left, nil
		}
	}
}

func (p *parser) parseTerm() (formulaValue, error) {
	left, err := p.parseFactor()
	if err != nil {
		return formulaValue{}, err
	}
	for {
		switch p.peek().kind {
		case tokenMul:
			p.next()
			right, err := p.parseFactor()
			if err != nil {
				return formulaValue{}, err
			}
			left, err = combineNumeric(left, right, func(a float64, b float64) float64 { return a * b })
			if err != nil {
				return formulaValue{}, err
			}
		case tokenDiv:
			p.next()
			right, err := p.parseFactor()
			if err != nil {
				return formulaValue{}, err
			}
			left, err = combineNumeric(left, right, func(a float64, b float64) float64 {
				if b == 0 {
					return math.Inf(1)
				}
				return a / b
			})
			if err != nil {
				return formulaValue{}, err
			}
		default:
			return left, nil
		}
	}
}

func (p *parser) parseFactor() (formulaValue, error) {
	tok := p.peek()
	switch tok.kind {
	case tokenPlus:
		p.next()
		return p.parseFactor()
	case tokenMinus:
		p.next()
		value, err := p.parseFactor()
		if err != nil {
			return formulaValue{}, err
		}
		number, err := value.asNumber()
		if err != nil {
			return formulaValue{}, err
		}
		return numberValue(-number), nil
	case tokenNumber:
		p.next()
		number, _ := strconv.ParseFloat(tok.text, 64)
		return numberValue(number), nil
	case tokenString:
		p.next()
		return stringValue(tok.text), nil
	case tokenLParen:
		p.next()
		value, err := p.parseComparison()
		if err != nil {
			return formulaValue{}, err
		}
		if _, err := p.expect(tokenRParen); err != nil {
			return formulaValue{}, err
		}
		return value, nil
	case tokenIdentifier:
		name := strings.ToUpper(tok.text)
		p.next()
		if p.peek().kind == tokenLParen {
			p.next()
			args, err := p.parseArguments()
			if err != nil {
				return formulaValue{}, err
			}
			return applyFunction(name, args)
		}
		switch name {
		case "TRUE":
			return boolValue(true), nil
		case "FALSE":
			return boolValue(false), nil
		default:
			return formulaValue{}, fmt.Errorf("unsupported identifier %s", name)
		}
	case tokenCell:
		start, err := workbook.ParseA1(tok.text)
		if err != nil {
			return formulaValue{}, err
		}
		p.next()
		if p.peek().kind == tokenColon {
			p.next()
			endTok, err := p.expect(tokenCell)
			if err != nil {
				return formulaValue{}, err
			}
			end, err := workbook.ParseA1(endTok.text)
			if err != nil {
				return formulaValue{}, err
			}
			if start.Row > end.Row {
				start.Row, end.Row = end.Row, start.Row
			}
			if start.Col > end.Col {
				start.Col, end.Col = end.Col, start.Col
			}
			values, err := p.eval.evalRange(start, end)
			if err != nil {
				return formulaValue{}, err
			}
			return rangeValue(values), nil
		}
		return p.eval.evalCellValue(start)
	default:
		return formulaValue{}, fmt.Errorf("unexpected token %q", tok.text)
	}
}

func (p *parser) parseArguments() ([]formulaValue, error) {
	args := []formulaValue{}
	if p.peek().kind == tokenRParen {
		p.next()
		return args, nil
	}
	for {
		value, err := p.parseExpressionOrRange()
		if err != nil {
			return nil, err
		}
		args = append(args, value)
		if p.peek().kind == tokenComma {
			p.next()
			continue
		}
		if _, err := p.expect(tokenRParen); err != nil {
			return nil, err
		}
		return args, nil
	}
}

func (p *parser) parseExpressionOrRange() (formulaValue, error) {
	if p.peek().kind == tokenCell && p.index+1 < len(p.tokens) && p.tokens[p.index+1].kind == tokenColon {
		return p.parseFactor()
	}
	return p.parseComparison()
}

func (p *parser) peek() token {
	if p.index >= len(p.tokens) {
		return token{kind: tokenEOF}
	}
	return p.tokens[p.index]
}

func (p *parser) next() token {
	tok := p.peek()
	p.index++
	return tok
}

func (p *parser) expect(kind tokenKind) (token, error) {
	tok := p.next()
	if tok.kind != kind {
		return token{}, fmt.Errorf("expected %v but found %q", kind, tok.text)
	}
	return tok, nil
}

func applyFunction(name string, args []formulaValue) (formulaValue, error) {
	flat := flattenArgs(args)
	switch name {
	case "SUM":
		values, _ := numericArgs(flat)
		sum := 0.0
		for _, value := range values {
			sum += value
		}
		return numberValue(sum), nil
	case "AVG", "AVERAGE":
		values, _ := numericArgs(flat)
		if len(values) == 0 {
			return numberValue(0), nil
		}
		sum := 0.0
		for _, value := range values {
			sum += value
		}
		return numberValue(sum / float64(len(values))), nil
	case "MIN":
		values, _ := numericArgs(flat)
		if len(values) == 0 {
			return numberValue(0), nil
		}
		min := values[0]
		for _, value := range values[1:] {
			if value < min {
				min = value
			}
		}
		return numberValue(min), nil
	case "MAX":
		values, _ := numericArgs(flat)
		if len(values) == 0 {
			return numberValue(0), nil
		}
		max := values[0]
		for _, value := range values[1:] {
			if value > max {
				max = value
			}
		}
		return numberValue(max), nil
	case "COUNT":
		count := 0
		for _, value := range flat {
			if _, err := value.asNumber(); err == nil {
				count++
			}
		}
		return numberValue(float64(count)), nil
	case "COUNTA":
		count := 0
		for _, value := range flat {
			if value.kind != valueBlank && strings.TrimSpace(value.outputString()) != "" {
				count++
			}
		}
		return numberValue(float64(count)), nil
	case "IF":
		if len(args) < 2 || len(args) > 3 {
			return formulaValue{}, fmt.Errorf("IF expects 2 or 3 arguments")
		}
		if args[0].truthy() {
			return args[1].scalar(), nil
		}
		if len(args) == 3 {
			return args[2].scalar(), nil
		}
		return blankValue(), nil
	case "ABS":
		value, err := requireArgCount(name, flat, 1)
		if err != nil {
			return formulaValue{}, err
		}
		number, err := value[0].asNumber()
		if err != nil {
			return formulaValue{}, err
		}
		return numberValue(math.Abs(number)), nil
	case "ROUND":
		if len(flat) < 1 || len(flat) > 2 {
			return formulaValue{}, fmt.Errorf("ROUND expects 1 or 2 arguments")
		}
		number, err := flat[0].asNumber()
		if err != nil {
			return formulaValue{}, err
		}
		digits := 0.0
		if len(flat) == 2 {
			digits, err = flat[1].asNumber()
			if err != nil {
				return formulaValue{}, err
			}
		}
		factor := math.Pow(10, math.Trunc(digits))
		if factor == 0 {
			return formulaValue{}, fmt.Errorf("invalid ROUND precision")
		}
		return numberValue(math.Round(number*factor) / factor), nil
	case "POWER":
		value, err := requireArgCount(name, flat, 2)
		if err != nil {
			return formulaValue{}, err
		}
		base, err := value[0].asNumber()
		if err != nil {
			return formulaValue{}, err
		}
		exponent, err := value[1].asNumber()
		if err != nil {
			return formulaValue{}, err
		}
		return numberValue(math.Pow(base, exponent)), nil
	case "MOD":
		value, err := requireArgCount(name, flat, 2)
		if err != nil {
			return formulaValue{}, err
		}
		left, err := value[0].asNumber()
		if err != nil {
			return formulaValue{}, err
		}
		right, err := value[1].asNumber()
		if err != nil {
			return formulaValue{}, err
		}
		if right == 0 {
			return formulaValue{}, fmt.Errorf("MOD division by zero")
		}
		return numberValue(math.Mod(left, right)), nil
	case "SQRT":
		value, err := requireArgCount(name, flat, 1)
		if err != nil {
			return formulaValue{}, err
		}
		number, err := value[0].asNumber()
		if err != nil {
			return formulaValue{}, err
		}
		if number < 0 {
			return formulaValue{}, fmt.Errorf("SQRT requires a non-negative number")
		}
		return numberValue(math.Sqrt(number)), nil
	case "AND":
		for _, value := range flat {
			if !value.truthy() {
				return boolValue(false), nil
			}
		}
		return boolValue(true), nil
	case "OR":
		for _, value := range flat {
			if value.truthy() {
				return boolValue(true), nil
			}
		}
		return boolValue(false), nil
	case "NOT":
		value, err := requireArgCount(name, flat, 1)
		if err != nil {
			return formulaValue{}, err
		}
		return boolValue(!value[0].truthy()), nil
	case "LEN":
		value, err := requireArgCount(name, flat, 1)
		if err != nil {
			return formulaValue{}, err
		}
		return numberValue(float64(utf8.RuneCountInString(value[0].outputString()))), nil
	case "CONCAT", "CONCATENATE":
		parts := make([]string, 0, len(flat))
		for _, value := range flat {
			parts = append(parts, value.outputString())
		}
		return stringValue(strings.Join(parts, "")), nil
	default:
		return formulaValue{}, fmt.Errorf("unsupported function %s", name)
	}
}

func flattenArgs(args []formulaValue) []formulaValue {
	flat := []formulaValue{}
	for _, arg := range args {
		if arg.kind == valueRange {
			flat = append(flat, arg.rangev...)
			continue
		}
		flat = append(flat, arg)
	}
	return flat
}

func numericArgs(args []formulaValue) ([]float64, error) {
	values := []float64{}
	for _, arg := range args {
		number, err := arg.asNumber()
		if err != nil {
			continue
		}
		values = append(values, number)
	}
	return values, nil
}

func requireArgCount(name string, args []formulaValue, count int) ([]formulaValue, error) {
	if len(args) != count {
		return nil, fmt.Errorf("%s expects %d arguments", name, count)
	}
	return args, nil
}

func combineNumeric(left formulaValue, right formulaValue, op func(float64, float64) float64) (formulaValue, error) {
	leftNumber, err := left.asNumber()
	if err != nil {
		return formulaValue{}, err
	}
	rightNumber, err := right.asNumber()
	if err != nil {
		return formulaValue{}, err
	}
	result := op(leftNumber, rightNumber)
	if math.IsInf(result, 0) || math.IsNaN(result) {
		return formulaValue{}, fmt.Errorf("invalid numeric result")
	}
	return numberValue(result), nil
}

func compareValues(left formulaValue, right formulaValue, op tokenKind) (bool, error) {
	leftNumber, leftErr := left.asNumber()
	rightNumber, rightErr := right.asNumber()
	if leftErr == nil && rightErr == nil {
		switch op {
		case tokenEq:
			return leftNumber == rightNumber, nil
		case tokenNE:
			return leftNumber != rightNumber, nil
		case tokenLT:
			return leftNumber < rightNumber, nil
		case tokenLTE:
			return leftNumber <= rightNumber, nil
		case tokenGT:
			return leftNumber > rightNumber, nil
		case tokenGTE:
			return leftNumber >= rightNumber, nil
		}
	}
	leftText := left.outputString()
	rightText := right.outputString()
	switch op {
	case tokenEq:
		return leftText == rightText, nil
	case tokenNE:
		return leftText != rightText, nil
	case tokenLT:
		return leftText < rightText, nil
	case tokenLTE:
		return leftText <= rightText, nil
	case tokenGT:
		return leftText > rightText, nil
	case tokenGTE:
		return leftText >= rightText, nil
	default:
		return false, fmt.Errorf("unsupported comparison")
	}
}

func tokenize(input string) []token {
	tokens := []token{}
	for index := 0; index < len(input); {
		ch := rune(input[index])
		if unicode.IsSpace(ch) {
			index++
			continue
		}
		switch ch {
		case '+':
			tokens = append(tokens, token{kind: tokenPlus, text: "+"})
			index++
		case '-':
			tokens = append(tokens, token{kind: tokenMinus, text: "-"})
			index++
		case '*':
			tokens = append(tokens, token{kind: tokenMul, text: "*"})
			index++
		case '/':
			tokens = append(tokens, token{kind: tokenDiv, text: "/"})
			index++
		case '(':
			tokens = append(tokens, token{kind: tokenLParen, text: "("})
			index++
		case ')':
			tokens = append(tokens, token{kind: tokenRParen, text: ")"})
			index++
		case ',':
			tokens = append(tokens, token{kind: tokenComma, text: ","})
			index++
		case ':':
			tokens = append(tokens, token{kind: tokenColon, text: ":"})
			index++
		case '=':
			tokens = append(tokens, token{kind: tokenEq, text: "="})
			index++
		case '<':
			if index+1 < len(input) && input[index+1] == '=' {
				tokens = append(tokens, token{kind: tokenLTE, text: "<="})
				index += 2
				continue
			}
			if index+1 < len(input) && input[index+1] == '>' {
				tokens = append(tokens, token{kind: tokenNE, text: "<>"})
				index += 2
				continue
			}
			tokens = append(tokens, token{kind: tokenLT, text: "<"})
			index++
		case '>':
			if index+1 < len(input) && input[index+1] == '=' {
				tokens = append(tokens, token{kind: tokenGTE, text: ">="})
				index += 2
				continue
			}
			tokens = append(tokens, token{kind: tokenGT, text: ">"})
			index++
		case '"':
			index++
			var builder strings.Builder
			for index < len(input) {
				if input[index] == '"' {
					if index+1 < len(input) && input[index+1] == '"' {
						builder.WriteByte('"')
						index += 2
						continue
					}
					index++
					break
				}
				builder.WriteByte(input[index])
				index++
			}
			tokens = append(tokens, token{kind: tokenString, text: builder.String()})
		default:
			if unicode.IsDigit(ch) || ch == '.' {
				start := index
				index++
				for index < len(input) {
					next := rune(input[index])
					if !unicode.IsDigit(next) && next != '.' {
						break
					}
					index++
				}
				tokens = append(tokens, token{kind: tokenNumber, text: input[start:index]})
				continue
			}
			if unicode.IsLetter(ch) || ch == '_' {
				start := index
				index++
				for index < len(input) {
					next := rune(input[index])
					if !unicode.IsLetter(next) && !unicode.IsDigit(next) && next != '_' {
						break
					}
					index++
				}
				text := input[start:index]
				kind := tokenIdentifier
				hasDigit := false
				hasLetter := false
				for _, r := range text {
					if unicode.IsLetter(r) {
						hasLetter = true
					}
					if unicode.IsDigit(r) {
						hasDigit = true
					}
				}
				if hasLetter && hasDigit {
					kind = tokenCell
				}
				tokens = append(tokens, token{kind: kind, text: text})
				continue
			}
			index++
		}
	}
	tokens = append(tokens, token{kind: tokenEOF})
	return tokens
}

func parseLiteralValue(value string) formulaValue {
	trimmed := strings.TrimSpace(value)
	if trimmed == "" {
		return blankValue()
	}
	if strings.EqualFold(trimmed, "true") {
		return boolValue(true)
	}
	if strings.EqualFold(trimmed, "false") {
		return boolValue(false)
	}
	number, err := strconv.ParseFloat(trimmed, 64)
	if err == nil {
		return numberValue(number)
	}
	return stringValue(value)
}

func (v formulaValue) asNumber() (float64, error) {
	switch v.kind {
	case valueNumber:
		return v.number, nil
	case valueBool:
		if v.boolv {
			return 1, nil
		}
		return 0, nil
	case valueString:
		number, err := strconv.ParseFloat(strings.TrimSpace(v.text), 64)
		if err != nil {
			return 0, fmt.Errorf("expected numeric value, got %q", v.text)
		}
		return number, nil
	case valueBlank:
		return 0, nil
	default:
		return 0, fmt.Errorf("expected numeric value")
	}
}

func (v formulaValue) truthy() bool {
	switch v.kind {
	case valueBool:
		return v.boolv
	case valueNumber:
		return v.number != 0
	case valueString:
		trimmed := strings.TrimSpace(v.text)
		if strings.EqualFold(trimmed, "true") {
			return true
		}
		if strings.EqualFold(trimmed, "false") || trimmed == "" {
			return false
		}
		number, err := strconv.ParseFloat(trimmed, 64)
		if err == nil {
			return number != 0
		}
		return true
	default:
		return false
	}
}

func (v formulaValue) scalar() formulaValue {
	if v.kind == valueRange {
		if len(v.rangev) == 0 {
			return blankValue()
		}
		return v.rangev[0]
	}
	return v
}

func (v formulaValue) outputString() string {
	switch v.kind {
	case valueNumber:
		return formatNumber(v.number)
	case valueString:
		return v.text
	case valueBool:
		if v.boolv {
			return "TRUE"
		}
		return "FALSE"
	default:
		return ""
	}
}

func blankValue() formulaValue {
	return formulaValue{kind: valueBlank}
}

func numberValue(value float64) formulaValue {
	return formulaValue{kind: valueNumber, number: value}
}

func stringValue(value string) formulaValue {
	return formulaValue{kind: valueString, text: value}
}

func boolValue(value bool) formulaValue {
	return formulaValue{kind: valueBool, boolv: value}
}

func rangeValue(values []formulaValue) formulaValue {
	return formulaValue{kind: valueRange, rangev: values}
}

func formatNumber(value float64) string {
	if math.Abs(value-math.Round(value)) < 1e-9 {
		return strconv.FormatInt(int64(math.Round(value)), 10)
	}
	return strconv.FormatFloat(value, 'f', -1, 64)
}
