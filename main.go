package main

import (
	"encoding/csv"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/lxn/walk"
	. "github.com/lxn/walk/declarative"
	"golang.org/x/text/encoding/japanese"
	"golang.org/x/text/transform"
)

var (
	rowTargets      = 1
	rowStartTask    = 2
	colStartTargets = 20
	colCyclic       = 3
	colPeriod       = 4
	colTitle        = 5
	colDetail       = 6

	iCyclic = 0
	iPeriod = 1
	iTitle  = 2
	iDetail = 3

	iTaskStartDay  = 3
	iTaskStartTime = 4
	iTaskEndtDay   = 5
	iTaskEndtTime  = 6
	iTaskTitle     = 8
	iTaskDetail    = 10

	weekday = map[string]int{"日": 0, "月": 1, "火": 2, "水": 3, "木": 4, "金": 6, "土": 6}

	start = time.Date(2021, 3, 13, 0, 0, 0, 0, time.Local)
	end   = time.Date(2022, 3, 13, 0, 0, 0, 0, time.Local)
)

type MyMainWindow struct {
	*walk.MainWindow
	comboTarget    *walk.ComboBox
	start          *walk.LineEdit
	end            *walk.LineEdit
	comboOuputType *walk.ComboBox
	sheet          [][]string
}

func main() {
	if err := _main(); err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
}

func _main() error {
	file := getExcelfile()
	if file == "" {
		return nil
	}

	// sheets := getSheets(f)
	sheet := getSheet(file, "Sheet1")
	// printCols2(sheet)
	// fmt.Println(getStartColumn())

	MW := getMainWindow(sheet)
	MW.Run()

	// fmt.Println(getPlainTasks(sheet, 2))
	return nil
}

func (mw *MyMainWindow) writeCsv() {
	i := mw.comboTarget.CurrentIndex()
	if i == -1 {
		walk.MsgBox(mw, "message", "担当者を選択してください", walk.MsgBoxOK|walk.MsgBoxIconError)
		return
	}
	writeCsv(convertTask(getPlainTasks(mw.sheet, i)))
}

func getExcelfile() string {
	var f string
	if len(os.Args) > 1 {
		f = os.Args[1]
	} else {
		files, err := filepath.Glob("./*.xlsx")
		if err != nil {
			log.Println(err)
			return ""
		}
		f = files[0]
	}
	info, err := os.Stat(f)
	if err != nil {
		log.Printf("Not exists '%s'\n", f)
		return ""
	}
	if info.IsDir() {
		log.Printf("'%s' is a directory\n", f)
		return ""
	}
	return f
}

func getMainWindow(sheet [][]string) MainWindow {
	mw := &MyMainWindow{}
	mw.sheet = sheet
	now := time.Now()
	textStart := now.Format("2006-01-02")
	end := now.AddDate(1, 0, 0)
	textEnd := end.Format("2006-01-02")
	MW := MainWindow{
		AssignTo: &mw.MainWindow,
		Title:    "generateSchedule",
		MinSize:  Size{Width: 250, Height: 300},
		Size:     Size{Width: 250, Height: 300},
		Layout:   VBox{},
		Children: []Widget{
			HSplitter{
				Children: []Widget{
					VSplitter{
						Children: []Widget{
							Label{Text: "担当者: "},
							VSpacer{},
						},
					},
					ComboBox{
						AssignTo: &mw.comboTarget,
						Model:    getTargets(mw.sheet),
					},
				},
			},
			HSplitter{
				Children: []Widget{
					VSplitter{
						Children: []Widget{
							Label{Text: "開始日: "},
							VSpacer{},
						},
					},
					VSplitter{
						Children: []Widget{
							LineEdit{
								AssignTo: &mw.start,
								Text:     textStart,
							},
							VSpacer{},
						},
					},
				},
			},
			HSplitter{
				Children: []Widget{
					VSplitter{
						Children: []Widget{
							Label{Text: "終了日: "},
							VSpacer{},
						},
					},
					VSplitter{
						Children: []Widget{
							LineEdit{
								AssignTo: &mw.end,
								Text:     textEnd,
							},
							VSpacer{},
						},
					},
				},
			},
			HSplitter{
				Children: []Widget{
					VSplitter{
						Children: []Widget{
							Label{Text: "出力形式: "},
							VSpacer{},
						},
					},
					ComboBox{
						AssignTo: &mw.comboOuputType,
						Model:    []string{"スケジュール", "タスク"},
					},
				},
			},
			PushButton{
				Text:      "Run",
				OnClicked: mw.writeCsv,
			},
		},
	}
	return MW
}

func getPicks(task []string) ([]time.Time, string) {
	var picks []time.Time
	var cycle string
	if task[iCyclic] == "年" {
		r := regexp.MustCompile(`\d+`)
		result := r.FindAllStringSubmatch(task[iPeriod], -1)
		for _, period := range result {
			month, err := strconv.Atoi(period[0])
			if err != nil {
				log.Println(err)
			}
			picks = append(picks, time.Date(start.Year(), time.Month(month), 1, 0, 0, 0, 0, time.Local))
		}
		cycle = "Y"
	} else if task[iCyclic] == "月" {
		r := regexp.MustCompile(`\d+`)
		result := r.FindAllStringSubmatch(task[iPeriod], -1)
		for _, period := range result {
			day, err := strconv.Atoi(period[0])
			if err != nil {
				log.Println(err)
			}
			picks = append(picks, time.Date(start.Year(), start.Month(), day, 0, 0, 0, 0, time.Local))
		}
		cycle = "M"
	} else if task[iCyclic] == "週" {
		var diff int
		r := regexp.MustCompile(`[月火水木金土日]`)
		result := r.FindAllStringSubmatch(task[iPeriod], -1)
		for _, period := range result {
			current_weekday := weekday[period[0]]
			if int(start.Weekday()) <= current_weekday {
				diff = current_weekday - int(start.Weekday())
			} else {
				diff = current_weekday - int(start.Weekday()) + 7
			}
			picks = append(picks, start.AddDate(0, 0, diff))
		}
		cycle = "W"
	} else {
		return nil, ""
	}
	return picks, cycle
}

func writeCsv(records [][]string) {
	f, err := os.Create("import.csv")
	if err != nil {
		log.Fatal(err)
	}
	w := csv.NewWriter(transform.NewWriter(f, japanese.ShiftJIS.NewEncoder()))
	w.UseCRLF = true
	w.WriteAll(records)
	w.Flush()

	if err := w.Error(); err != nil {
		log.Fatal(err)
	}
}

func convertTask(ts [][]string) [][]string {
	var tasks [][]string
	var header = []string{"ユーザー/グループシステムID", "氏名/グループ名", "ＩＤ（システムＩＤ：自動発番）", "開始日", "開始時刻", "終了日", "終了時刻", "予定", "件名", "場所", "内容", "プライベート", "外出区分", "重要度", "予約種別", "帯状", "フラグ", "アイコン番号", "承認依頼", "確認通知メール"}
	tasks = append(tasks, header)

	for _, t := range ts {
		picks, cycle := getPicks(t)
		if cycle == "" {
			continue
		}
		for _, pick := range picks {
			for pick.Unix() <= end.Unix() {
				task := make([]string, 20, 20)
				task[iTaskStartDay] = pick.Format("2006-01-02")
				task[iTaskStartTime] = "8:30"
				task[iTaskEndtDay] = pick.Format("2006-01-02")
				task[iTaskEndtTime] = "8:30"
				task[iTaskTitle] = t[iTitle]
				task[iTaskDetail] = t[iDetail]
				tasks = append(tasks, task)
				if cycle == "Y" {
					pick = pick.AddDate(1, 0, 0)
				} else if cycle == "M" {
					pick = pick.AddDate(0, 1, 0)
				} else {
					// cycle == "W"
					pick = pick.AddDate(0, 0, 7)
				}
			}
		}

	}
	return tasks
}

func getPlainTasks(s [][]string, indexTarget int) [][]string {
	var tasks [][]string
	var task []string
	colTarget := colStartTargets + indexTarget

	for i, cells := range s {
		if i < rowStartTask {
			continue
		}
		// fmt.Println(cols[colTarget])
		if cells[colTarget] != "〇" {
			continue
		}
		task = []string{cells[colCyclic], cells[colPeriod], cells[colTitle], cells[colDetail]}
		tasks = append(tasks, task)
	}
	return tasks
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
