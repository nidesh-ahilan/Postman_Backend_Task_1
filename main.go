package main

import (
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {

	if len(os.Args) < 2 {
		fmt.Println("Usage: go run main.go <path-to-xlsx-file>")
		return
	}

	file := os.Args[1]

	f, err := excelize.OpenFile(file)
	if err != nil {
		fmt.Println(err)
		return
	}

	sheetName := f.GetSheetList()[0]

	rows, err := f.GetRows(sheetName)
	if err != nil {
		fmt.Println("Error reading rows:", err)
		return
	}

	for index, data := range rows {
		if index == 0 {
			continue
		}

		if isEmptyRow(data) {
			continue
		}

		quiz, _ := strconv.ParseFloat(strings.TrimSpace(data[3]), 32)
		midSem, _ := strconv.ParseFloat(strings.TrimSpace(data[4]), 32)
		labTest, _ := strconv.ParseFloat(strings.TrimSpace(data[5]), 32)
		weeklyLabs, _ := strconv.ParseFloat(strings.TrimSpace(data[6]), 32)
		compre, _ := strconv.ParseFloat(strings.TrimSpace(data[8]), 32)

		calculatedPreCompre := (quiz + midSem + labTest + weeklyLabs)
		calculatedCompre := (calculatedPreCompre + compre)

		elementPreCompre := fmt.Sprintf("L%d", index+1)
		elementCompre := fmt.Sprintf("K%d", index+1)

		f.SetCellValue(sheetName, elementPreCompre, fmt.Sprintf("%.2f", calculatedPreCompre))
		f.SetCellValue(sheetName, elementCompre, fmt.Sprintf("%.2f", calculatedCompre))
	}

	for _, data := range rows {
		if isEmptyRow(data) {
			continue
		}

		fmt.Println(strings.Join(data, "\t"))
	}

	if err := f.Save(); err != nil {
		fmt.Println("Error saving the file:", err)
	}

}

func isEmptyRow(row []string) bool {
	for _, cell := range row {
		if strings.TrimSpace(cell) != "" {
			return false
		}
	}
	return true
}
