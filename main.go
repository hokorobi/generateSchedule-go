package main

import (
	"fmt"
	"log"
	"os"

	"github.com/360EntSecGroup-Skylar/excelize"
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
		fmt.Printf("Not exists '%s'\n", f)
		return nil
	}
	if info.IsDir() {
		return nil
	}
	sheets := getSheets(f)
	printCols(sheets)
	// fmt.Println(getStartColumn())
	return nil
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

func printCols(sheets map[string][][]string) {
	// O_WRONLY:書き込みモード開く, O_CREATE:無かったらファイルを作成
	for _, rows := range sheets {
		for _, cols := range rows {
			fmt.Println(cols)
		}
	}
}

// vim: ff=unix
