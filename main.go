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

	f.SetCellValue(sheetName, "L1", "Calculated Total")

	for index, data := range rows {
		if index == 0 {
			continue
		}

		if isEmptyRow(data) {
			continue
		}

		quiz, _ := strconv.ParseFloat(strings.TrimSpace(data[4]), 32)
		midSem, _ := strconv.ParseFloat(strings.TrimSpace(data[5]), 32)
		labTest, _ := strconv.ParseFloat(strings.TrimSpace(data[6]), 32)
		weeklyLabs, _ := strconv.ParseFloat(strings.TrimSpace(data[7]), 32)
		compre, _ := strconv.ParseFloat(strings.TrimSpace(data[9]), 32)
		givenTotal, _ := strconv.ParseFloat(strings.TrimSpace(data[10]), 32)

		calculatedPreCompre := (quiz + midSem + labTest + weeklyLabs)
		calculatedTotal := (calculatedPreCompre + compre)

		elementCompre := fmt.Sprintf("L%d", index+1)

		f.SetCellValue(sheetName, elementCompre, fmt.Sprintf("%.2f", calculatedTotal))

		if givenTotal-calculatedTotal > 0.02 {
			fmt.Printf("Total not matching in row %d as given total is %.2f but the actual total is %.2f\n", index+1, givenTotal, calculatedTotal)
		}

	}

	rows, _ = f.GetRows(sheetName)

	fmt.Println("\n")

	average(rows, 4, "Quiz")
	average(rows, 5, "Mid Sem")
	average(rows, 6, "Lab Test")
	average(rows, 7, "Weekly Labs")
	average(rows, 9, "Compre")
	average(rows, 11, "Total")

	fmt.Printf("\n")

	fmt.Println("Branch wise average (Only 2024 Batch)")

	branchAverage(rows, 11, "Total", "2024A3", "EEE")
	branchAverage(rows, 11, "Total", "2024A4", "Mechanical")
	branchAverage(rows, 11, "Total", "2024A5", "B. Pharma")
	branchAverage(rows, 11, "Total", "2024A7", "CSE")
	branchAverage(rows, 11, "Total", "2024A8", "ENI")
	branchAverage(rows, 11, "Total", "2024AA", "ECE")
	branchAverage(rows, 11, "Total", "2024AD", "MNC")

	fmt.Printf("\n")

	findTopThree(rows)

	f.SaveAs(file)

}

func isEmptyRow(row []string) bool {
	for _, cell := range row {
		if strings.TrimSpace(cell) != "" {
			return false
		}
	}
	return true
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

func branchAverage(rows [][]string, columnIndex int, columnName string, prefix string, branch string) {
	var sum float64
	var count float64

	count = 0
	sum = 0

	for _, data := range rows {
		if strings.HasPrefix(data[3], prefix) {
			value, _ := strconv.ParseFloat(strings.TrimSpace(data[columnIndex]), 64)
			sum += value
			count++
		}
	}

	if count > 0 {
		avg := sum / count
		fmt.Printf("Average of %s: %.2f\n", branch, avg)
	}
}

func findTopThree(rows [][]string) {
	first := 1
	second := 2
	third := 3

	for i := 3; i < len(rows); i++ {
		total, _ := strconv.ParseFloat(rows[i][11], 64)
		firstTotal, _ := strconv.ParseFloat(rows[first][11], 64)
		secondTotal, _ := strconv.ParseFloat(rows[second][11], 64)
		thirdTotal, _ := strconv.ParseFloat(rows[third][11], 64)

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

	fmt.Printf("Rank 1 : ID : %s , Marks: %s \n", rows[first][2], rows[first][11])
	fmt.Printf("Rank 2 : ID : %s , Marks: %s \n", rows[second][2], rows[second][11])
	fmt.Printf("Rank 3 : ID : %s , Marks: %s \n", rows[third][2], rows[third][11])
}
