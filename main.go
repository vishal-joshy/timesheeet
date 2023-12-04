package main

import (
	"encoding/csv"
	"fmt"
	"os"
	"strconv"

	"github.com/xuri/excelize/v2"
)

type Task struct {
	date            string
	task            string
	taskDescription string
	businessFunc    string
	hours           int
}

func ReadCSV() [][]string {
	f, err := os.Open("task.csv")
	if err != nil {
		panic(err)
	}
	defer f.Close()

	csvReader := csv.NewReader(f)
	data, err := csvReader.ReadAll()
	if err != nil {
		panic(err)
	}
	return data
}

func writeTasks(username string, tasks []Task) {
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	_, err := f.NewSheet("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}
	f.SetCellValue("Sheet1", "A1", "CONSULTANT WEEKLY TIMESHEET")

	err = f.MergeCell("Sheet1", "A1", "C1")
	if err != nil {
		panic(err)
	}

	f.SetCellValue("Sheet1", "A3", "Consultant Name")
	err = f.MergeCell("Sheet1", "A3", "B3")
	if err != nil {
		panic(err)
	}
	f.SetCellValue("Sheet1", "C3", username)
	f.SetCellValue("Sheet1", "A4", "Start Date")
	f.SetCellValue("Sheet1", "C4", "End Date")
	f.SetCellValue("Sheet1", "A5", tasks[0].date)
	f.SetCellValue("Sheet1", "C5", tasks[len(tasks)-1].date)

	f.SetSheetRow("Sheet1", "B7", &[]interface{}{"Date", "Task", "Task Name", "Business Function", "Task Description", "Notes", "Hours", "Total Time", "Total Hours per Day"})

	totalHours := 0
	i := 0

	for _, v := range tasks {
		f.SetSheetRow("Sheet1", fmt.Sprintf("B%v", i+8), &[]interface{}{v.date, i + 1, v.task, v.businessFunc, v.taskDescription, nil, v.hours, v.hours, v.hours})
		i++
		totalHours = totalHours + v.hours
	}

	f.SetCellValue("Sheet1", fmt.Sprintf("I%v", i+10), "Total Hours Per week")
	f.SetCellValue("Sheet1", fmt.Sprintf("J%v", i+10), totalHours)

	titleStyle, _ := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Color: []string{"204559"}, Pattern: 1},
		Font: &excelize.Font{
			Color: "FFFFFF",
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "up", Color: "000000", Style: 1},
			{Type: "down", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	totalBarStyle, _ := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Color: []string{"316886"}, Pattern: 1},
		Font: &excelize.Font{
			Color: "FFFFFF",
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "up", Color: "000000", Style: 1},
			{Type: "down", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	borderStyle, _ := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "up", Color: "000000", Style: 1},
			{Type: "down", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})

	_ = f.SetRowStyle("Sheet1", 1, 220, borderStyle)
	_ = f.SetColStyle("Sheet1", "A:N", borderStyle)
	_ = f.SetCellStyle("Sheet1", "A6", "N7", titleStyle)

	_ = f.SetCellStyle("Sheet1", fmt.Sprintf("A%v", i+10), fmt.Sprintf("N%v", i+10), totalBarStyle)

	if err := f.SaveAs("TimeSheet.xlsx"); err != nil {
		fmt.Println(err)
	}
}

func getTaskList(td [][]string) []Task {
	taskSlice := []Task{}
	for _, v := range td {
		task := Task{}
		task.date = v[0]
		task.task = v[1]
		task.businessFunc = v[2]
		task.taskDescription = v[3]
		hours, err := strconv.Atoi(v[4])
		if err != nil {
			panic(err)
		}
		task.hours = hours
		taskSlice = append(taskSlice, task)
	}
	return taskSlice
}

func main() {
	username := "Vishal Joshy"
	taskListData := ReadCSV()
	taskList := getTaskList(taskListData)
	fmt.Println(taskList)
	writeTasks(username, taskList)
}
