package main

import (
	"fmt"
	"log"
	"os"

	"github.com/360EntSecGroup-Skylar/excelize"
)

var (
	rowTargets      = 1
	rowStartTask    = 2
	colStartTargets = 20
	colCyclic       = 3
)

func main() {
	if err := _main(); err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
}

func _main() error {
	f := os.Args[1]
	info, err := os.Stat(f)
	if err != nil {
		log.Printf("Not exists '%s'\n", f)
		return nil
	}
	if info.IsDir() {
		return nil
	}
	// sheets := getSheets(f)
	sheet := getSheet(f, "Sheet1")
	// printCols2(sheet)
	// fmt.Println(getStartColumn())

	targets := getTargets(sheet)
	log.Println(targets)

	fmt.Println(getTargetSchedule(sheet, 1))
	return nil
}

func getTargetSchedule(s [][]string, indexTarget int) [][]string {
	var schedules [][]string
	var schedule []string
	colTarget := colStartTargets + indexTarget

	for i, cols := range s {
		if i < rowStartTask {
			continue
		}
		// fmt.Println(cols[colTarget])
		if cols[colTarget] != "〇" {
			continue
		}
		for j, cell := range cols {
			if j == colCyclic {
				schedule = []string{cell}
			}
		}
		schedules = append(schedules, schedule)
	}
	return schedules
}

// 担当者の配列を取得
func getTargets(s [][]string) []string {
	var targets []string
	for i, cols := range s {
		if i != rowTargets {
			continue
		}
		for j, cell := range cols {
			if j < colStartTargets {
				continue
			}
			targets = append(targets, cell)
		}
	}
	return targets
}

func getSheets(f string) map[string][][]string {
	sheets := map[string][][]string{}
	xls, err := excelize.OpenFile(f)
	if err != nil {
		log.Println(err)
		return sheets
	}
	for _, sheetName := range xls.GetSheetMap() {
		rows := xls.GetRows(sheetName)
		_, ok := sheets[sheetName]
		if ok {
			sheets[sheetName] = append(sheets[sheetName], rows...)
		} else {
			sheets[sheetName] = rows
		}
	}
	return sheets
}

func getSheet(f string, sheetName string) [][]string {
	xls, err := excelize.OpenFile(f)
	if err != nil {
		log.Println(err)
		return nil
	}
	return xls.GetRows(sheetName)
}

func printCols(sheets map[string][][]string) {
	for _, rows := range sheets {
		for _, cols := range rows {
			fmt.Println(cols)
		}
	}
}

func printCols2(sheet [][]string) {
	for _, cols := range sheet {
		fmt.Println(cols)
	}
}

// vim: ff=unix
