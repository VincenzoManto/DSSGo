package main

import (
	"errors"
	"flag"
	"fmt"
	"os"
	"strings"

	"github.com/vincm/dss/dss-go/internal/dashboard"
	"github.com/vincm/dss/dss-go/internal/dss"
	"github.com/vincm/dss/dss-go/internal/formula"
	"github.com/vincm/dss/dss-go/internal/importers"
	"github.com/vincm/dss/dss-go/internal/tui"
	"github.com/vincm/dss/dss-go/internal/workbook"
)

func main() {
	if err := run(os.Args[1:]); err != nil {
		fmt.Fprintln(os.Stderr, "error:", err)
		os.Exit(1)
	}
}

func run(args []string) error {
	if len(args) == 0 {
		printUsage()
		return nil
	}
	if tui.FileExists(args[0]) {
		return runTUI(args)
	}

	switch args[0] {
	case "convert":
		return runConvert(args[1:])
	case "set":
		return runSet(args[1:])
	case "tui", "open":
		return runTUI(args[1:])
	case "dashboard":
		return runDashboard(args[1:])
	case "help", "-h", "--help":
		printUsage()
		return nil
	default:
		return fmt.Errorf("unknown command %q", args[0])
	}
}

func printUsage() {
	fmt.Println("dss-go: import, inspect, edit, and export spreadsheet data as DSS")
	fmt.Println()
	fmt.Println("Usage:")
	fmt.Println("  dss <input> [--output output.dss]")
	fmt.Println("  dss tui <input> [--output output.dss]")
	fmt.Println("  dss dashboard <input>")
	fmt.Println("  dss convert <input> <output.dss>")
	fmt.Println("  dss set <input> <output.dss> --sheet NAME --cell A1 --value VALUE")
	fmt.Println()
	fmt.Println("Supported input formats: .dss, .csv, .xlsx, .xlsm, .xls")
	fmt.Println("Notes: formula evaluation supports arithmetic, comparisons, IF, logical and common math functions, plus cached imported Excel values for unsupported formulas.")
}

func runTUI(args []string) error {
	flagArgs, positionalArgs, err := splitFlagArgs(args, map[string]bool{"--output": true})
	if err != nil {
		return err
	}
	fs := flag.NewFlagSet("tui", flag.ContinueOnError)
	fs.SetOutput(os.Stderr)
	var outputPath string
	fs.StringVar(&outputPath, "output", "", "path to write DSS saves to")
	if err := fs.Parse(flagArgs); err != nil {
		return err
	}
	if len(positionalArgs) != 1 {
		return errors.New("tui requires exactly one input file")
	}
	return tui.Run(positionalArgs[0], outputPath)
}

func runDashboard(args []string) error {
	if len(args) != 1 {
		return errors.New("dashboard requires exactly one input file")
	}
	return dashboard.Run(args[0])
}

func runConvert(args []string) error {
	if len(args) != 2 {
		return errors.New("convert requires an input file and an output .dss path")
	}
	wb, err := importers.LoadWorkbook(args[0])
	if err != nil {
		return err
	}
	formula.RecalculateWorkbook(wb)
	return dss.WriteFile(args[1], wb)
}

func runSet(args []string) error {
	flagArgs, positionalArgs, err := splitFlagArgs(args, map[string]bool{"--sheet": true, "--cell": true, "--value": true})
	if err != nil {
		return err
	}
	fs := flag.NewFlagSet("set", flag.ContinueOnError)
	fs.SetOutput(os.Stderr)
	var sheetName string
	var cellRef string
	var value string
	fs.StringVar(&sheetName, "sheet", "", "sheet name to update")
	fs.StringVar(&cellRef, "cell", "", "A1 cell reference")
	fs.StringVar(&value, "value", "", "raw cell input, for example 123 or =SUM(A1:A3)")
	if err := fs.Parse(flagArgs); err != nil {
		return err
	}
	if len(positionalArgs) != 2 {
		return errors.New("set requires an input file and an output .dss path")
	}
	if sheetName == "" || cellRef == "" {
		return errors.New("set requires --sheet and --cell")
	}

	wb, err := importers.LoadWorkbook(positionalArgs[0])
	if err != nil {
		return err
	}
	sheet := wb.EnsureSheet(sheetName)
	coord, err := workbook.ParseA1(cellRef)
	if err != nil {
		return err
	}
	sheet.SetCell(coord.Row, coord.Col, value)
	formula.RecalculateWorkbook(wb)
	return dss.WriteFile(positionalArgs[1], wb)
}

func splitFlagArgs(args []string, flagsWithValues map[string]bool) ([]string, []string, error) {
	flagArgs := make([]string, 0, len(args))
	positionals := make([]string, 0, len(args))
	for index := 0; index < len(args); index++ {
		arg := args[index]
		if !strings.HasPrefix(arg, "-") || arg == "-" {
			positionals = append(positionals, arg)
			continue
		}
		flagArgs = append(flagArgs, arg)
		flagName := arg
		if equalsIndex := strings.Index(flagName, "="); equalsIndex >= 0 {
			flagName = flagName[:equalsIndex]
		}
		if flagsWithValues[flagName] && !strings.Contains(arg, "=") {
			if index+1 >= len(args) {
				return nil, nil, fmt.Errorf("flag %s requires a value", flagName)
			}
			index++
			flagArgs = append(flagArgs, args[index])
		}
	}
	return flagArgs, positionals, nil
}
