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
	f.SetCellValue("Sheet1", "A3", "Consultant Name")
	f.SetCellValue("Sheet1", "C3", username)
	f.SetCellValue("Sheet1", "A4", "Start Date")
	f.SetCellValue("Sheet1", "C4", "End Date")
	f.SetCellValue("Sheet1", "A5", tasks[0].date)
	f.SetCellValue("Sheet1", "C5", tasks[len(tasks)-1].date)

	f.SetCellValue("Sheet1", "B7", "Date")
	f.SetCellValue("Sheet1", "C7", "Task")
	f.SetCellValue("Sheet1", "D7", "Task Name")
	f.SetCellValue("Sheet1", "E7", "Business Function")
	f.SetCellValue("Sheet1", "F7", "Task Description")
	f.SetCellValue("Sheet1", "G7", "Notes")
	f.SetCellValue("Sheet1", "H7", "Hours")
	f.SetCellValue("Sheet1", "I7", "Total Time")
	f.SetCellValue("Sheet1", "J7", "Total Hours per day")

	totalHours := 0
	i := 0

	for _, v := range tasks {
		f.SetCellValue("Sheet1", fmt.Sprintf("B%v", i+8), v.date)
		f.SetCellValue("Sheet1", fmt.Sprintf("C%v", i+8), i+1)
		f.SetCellValue("Sheet1", fmt.Sprintf("D%v", i+8), v.task)
		f.SetCellValue("Sheet1", fmt.Sprintf("E%v", i+8), v.businessFunc)
		f.SetCellValue("Sheet1", fmt.Sprintf("F%v", i+8), v.taskDescription)
		f.SetCellValue("Sheet1", fmt.Sprintf("G%v", i+8), "")
		f.SetCellValue("Sheet1", fmt.Sprintf("H%v", i+8), v.hours)
		f.SetCellValue("Sheet1", fmt.Sprintf("I%v", i+8), v.hours)
		f.SetCellValue("Sheet1", fmt.Sprintf("J%v", i+8), v.hours)
		i++
		totalHours = totalHours + v.hours
	}

	f.SetCellValue("Sheet1", fmt.Sprintf("I%v", i+10), "Total Hours Per week")
	f.SetCellValue("Sheet1", fmt.Sprintf("J%v", i+10), totalHours)

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
