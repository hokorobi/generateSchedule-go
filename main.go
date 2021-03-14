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

const (
	rowStaff     = 1
	rowStartTask = 2

	colCyclic     = 3
	colPeriod     = 4
	colTitle      = 5
	colDetail     = 6
	colStartStaff = 20

	iCyclic = 0
	iPeriod = 1
	iTitle  = 2
	iDetail = 3

	iTaskTitle    = 0
	iTaskStartDay = 1
	iTaskEndtDay  = 2
	iTaskDetail   = 12

	iScheduleStartDay  = 3
	iScheduleStartTime = 4
	iScheduleEndtDay   = 5
	iScheduleEndtTime  = 6
	iScheduleTitle     = 8
	iScheduleDetail    = 10

	formatSchedule = 0
	formatTask     = 1
)

var (
	weekday = map[string]int{"日": 0, "月": 1, "火": 2, "水": 3, "木": 4, "金": 6, "土": 6}

	start = time.Date(2021, 3, 13, 0, 0, 0, 0, time.Local)
	end   = time.Date(2022, 3, 13, 0, 0, 0, 0, time.Local)

	filename  string
	sheetname string

	scheduleHeader = []string{
		"ユーザー/グループシステムID",
		"氏名/グループ名",
		"ＩＤ（システムＩＤ：自動発番）",
		"開始日",
		"開始時刻",
		"終了日",
		"終了時刻",
		"予定",
		"件名",
		"場所",
		"内容",
		"プライベート",
		"外出区分",
		"重要度",
		"予約種別",
		"帯状",
		"フラグ",
		"アイコン番号",
		"承認依頼",
		"確認通知メール",
	}

	taskHeader = []string{
		"件名",
		"開始日",
		"期限",
		"アラーム オン/オフ",
		"アラーム日付",
		"アラーム時刻",
		"完了日",
		"達成率",
		"予測時間",
		"実働時間",
		"Schedule+ の優先度",
		"プライベート",
		"メモ",
		"会社名",
		"経費情報",
		"支払い条件",
		"状況",
		"秘密度",
		"分類",
		"役割",
		"優先度",
		"連絡先",
	}
)

type MyMainWindow struct {
	*walk.MainWindow
	comboStaff     *walk.ComboBox
	start          *walk.LineEdit
	end            *walk.LineEdit
	comboOuputType *walk.ComboBox
	sheet          [][]string
}

func main() {
	sheet := getSheet()
	if sheet == nil {
		return
	}

	MW := getMainWindow(sheet)
	MW.Run()
}

func (mw *MyMainWindow) writeCsv() {
	staff := mw.comboStaff.CurrentIndex()
	if staff == -1 {
		walk.MsgBox(mw, "message", "担当者を選択してください", walk.MsgBoxOK|walk.MsgBoxIconError)
		return
	}
	formatType := mw.comboOuputType.CurrentIndex()
	if formatType == -1 {
		walk.MsgBox(mw, "message", "出力形式を選択してください", walk.MsgBoxOK|walk.MsgBoxIconError)
		return
	}
	// TODO: 開始日、終了日を取得
	writeCsv(convertOutlookFormat(getPlainTasks(mw.sheet, staff), formatType))
}

func convertOutlookFormat(tasks [][]string, formatType int) [][]string {
	if formatType == formatSchedule {
		return convertSchedule(tasks)
	} else {
		return convertTask(tasks)
	}
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
		MinSize:  Size{Width: 250, Height: 260},
		Size:     Size{Width: 250, Height: 260},
		Layout:   VBox{},
		Children: []Widget{
			Composite{
				Layout: Grid{Columns: 2},
				Children: []Widget{
					Label{Text: "担当者: "},
					ComboBox{
						AssignTo:     &mw.comboStaff,
						Model:        getStaves(mw.sheet),
						CurrentIndex: 0,
					},
					Label{Text: "開始日: "},
					LineEdit{
						AssignTo: &mw.start,
						Text:     textStart,
					},
					Label{Text: "終了日: "},
					LineEdit{
						AssignTo: &mw.end,
						Text:     textEnd,
					},
					Label{Text: "出力形式: "},
					ComboBox{
						AssignTo:     &mw.comboOuputType,
						Model:        []string{"スケジュール", "タスク"},
						CurrentIndex: 0,
					},
					PushButton{
						ColumnSpan: 2,
						Text:       "出力",
						OnClicked:  mw.writeCsv,
					},
					Label{Text: "ファイル名: "},
					Label{Text: filename},
					Label{Text: "シート名: "},
					Label{Text: sheetname},
					VSpacer{},
				},
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

func convertSchedule(tasks [][]string) [][]string {
	var lists [][]string
	lists = append(lists, scheduleHeader)

	for _, task := range tasks {
		picks, cycle := getPicks(task)
		if cycle == "" {
			continue
		}
		for _, pick := range picks {
			for pick.Unix() <= end.Unix() {
				list := make([]string, 20, 20)
				list[iScheduleStartDay] = pick.Format("2006-01-02")
				list[iScheduleStartTime] = "8:30"
				list[iScheduleEndtDay] = pick.Format("2006-01-02")
				list[iScheduleEndtTime] = "8:30"
				list[iScheduleTitle] = task[iTitle]
				list[iScheduleDetail] = task[iDetail]
				lists = append(lists, list)
				pick = getNextPick(pick, cycle)
			}
		}
	}
	return lists
}

func convertTask(tasks [][]string) [][]string {
	var lists [][]string
	lists = append(lists, taskHeader)

	for _, task := range tasks {
		picks, cycle := getPicks(task)
		if cycle == "" {
			continue
		}
		for _, pick := range picks {
			for pick.Unix() <= end.Unix() {
				list := make([]string, 22, 22)
				list[iTaskTitle] = task[iTitle]
				list[iTaskStartDay] = pick.Format("2006-01-02")
				list[iTaskEndtDay] = pick.Format("2006-01-02")
				list[iTaskDetail] = task[iDetail]
				lists = append(lists, list)
				pick = getNextPick(pick, cycle)
			}
		}
	}
	return lists
}

func getNextPick(pick time.Time, cycle string) time.Time {
	if cycle == "Y" {
		pick = pick.AddDate(1, 0, 0)
	} else if cycle == "M" {
		pick = pick.AddDate(0, 1, 0)
	} else {
		// cycle == "W"
		pick = pick.AddDate(0, 0, 7)
	}
	return pick
}
func getPlainTasks(s [][]string, indexStaff int) [][]string {
	var tasks [][]string
	var task []string
	colStaff := colStartStaff + indexStaff

	for i, cells := range s {
		if i < rowStartTask {
			continue
		}
		// fmt.Println(cols[colStaff])
		if cells[colStaff] != "〇" {
			continue
		}
		task = []string{cells[colCyclic], cells[colPeriod], cells[colTitle], cells[colDetail]}
		tasks = append(tasks, task)
	}
	return tasks
}

// 担当者の配列を取得
func getStaves(s [][]string) []string {
	var staves []string
	for i, cols := range s {
		if i != rowStaff {
			continue
		}
		for j, cell := range cols {
			if j < colStartStaff {
				continue
			}
			staves = append(staves, cell)
		}
	}
	return staves
}

func getSheetName(filename string) string {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		log.Println(err)
		return ""
	}
	return f.GetSheetMap()[1]
}

func getSheet() [][]string {
	filename = getExcelfile()
	if filename == "" {
		log.Println("Excel ファイルは見つかりませんでした。")
		return nil
	}

	sheetname = getSheetName(filename)
	if sheetname == "" {
		log.Printf("%s からシート名が取得できませんでした。", filename)
		return nil
	}

	f, err := excelize.OpenFile(filename)
	if err != nil {
		log.Println(err)
		return nil
	}
	return f.GetRows(sheetname)
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
