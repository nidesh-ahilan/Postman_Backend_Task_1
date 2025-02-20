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
		fmt.Println("File not found")
		return
	}

	file := os.Args[1]

	f, _ := excelize.OpenFile(file)

	sheetName := f.GetSheetList()[0]

	rows, _ := f.GetRows(sheetName)

	f.SetCellValue(sheetName, "K1", "Calculated Compre")

	for index, data := range rows {
		if index == 0 {
			continue
		}

		if isEmptyRow(data) {
			f.RemoveRow(sheetName, index+1)
			continue
		}

		quiz, _ := strconv.ParseFloat(strings.TrimSpace(data[3]), 32)
		midSem, _ := strconv.ParseFloat(strings.TrimSpace(data[4]), 32)
		labTest, _ := strconv.ParseFloat(strings.TrimSpace(data[5]), 32)
		weeklyLabs, _ := strconv.ParseFloat(strings.TrimSpace(data[6]), 32)
		compre, _ := strconv.ParseFloat(strings.TrimSpace(data[8]), 32)
		givenTotal, _ := strconv.ParseFloat(strings.TrimSpace(data[9]), 32)

		calculatedPreCompre := (quiz + midSem + labTest + weeklyLabs)
		calculatedTotal := (calculatedPreCompre + compre)

		elementCompre := fmt.Sprintf("K%d", index+1)

		f.SetCellValue(sheetName, elementCompre, fmt.Sprintf("%.2f", calculatedTotal))

		if givenTotal-calculatedTotal > 0.02 {
			fmt.Printf("Total not matching in row %d as given total is %.2f but the actual total is %.2f\n", index+1, givenTotal, calculatedTotal)
		}

	}

	fmt.Println("\n")

	average(rows, 3, "Quiz")
	average(rows, 4, "Mid Sem")
	average(rows, 5, "Lab Test")
	average(rows, 6, "Weekly Labs")
	average(rows, 8, "Compre")
	average(rows, 10, "Total")

	fmt.Printf("\n")

	findTopThree(rows)

	f.SaveAs(file)

}

func isEmptyRow(row []string) bool {
	for _, cell := range row {
		if strings.TrimSpace(cell) == "" {
			return true
		}
	}
	return false
}

func average(rows [][]string, columnIndex int, columnName string) {
	var sum float64
	var count float64

	count = -1
	sum = 0

	for _, data := range rows {

		value, _ := strconv.ParseFloat(strings.TrimSpace(data[columnIndex]), 64)
		sum += value
		count++

	}

	if count > 0 {
		avg := sum / count
		fmt.Printf("Average of %s: %.2f\n", columnName, avg)
	}
}

func findTopThree(rows [][]string) {
	first := 1
	second := 2
	third := 3

	for i := 3; i < len(rows); i++ {
		total, _ := strconv.ParseFloat(rows[i][10], 64)
		firstTotal, _ := strconv.ParseFloat(rows[first][10], 64)
		secondTotal, _ := strconv.ParseFloat(rows[second][10], 64)
		thirdTotal, _ := strconv.ParseFloat(rows[third][10], 64)

		if total > firstTotal {
			third = second
			second = first
			first = i
		} else if total > secondTotal {
			third = second
			second = i
		} else if total > thirdTotal {
			third = i
		}

	}

	fmt.Printf("Rank 1 : ID : %s , Marks: %s \n", rows[first][2], rows[first][10])
	fmt.Printf("Rank 2 : ID : %s , Marks: %s \n", rows[second][2], rows[second][10])
	fmt.Printf("Rank 3 : ID : %s , Marks: %s \n", rows[third][2], rows[third][10])
}
