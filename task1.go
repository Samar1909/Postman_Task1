package main

import (
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func get24batchSlice(f *excelize.File) []int {
	var myslice []int
	for i := 1; ; i++ {
		cellString := fmt.Sprintf("D%d", i)
		campusID, err := f.GetCellValue("CSF111_202425_01_GradeBook", cellString)
		if err != nil {
			fmt.Println("Error getting cell value for ", cellString)
		}
		if strings.Contains(campusID, "2024") {
			myslice = append(myslice, i)
		}
		if campusID == "" {
			break
		}
	}
	return myslice
}

func main() {
	if len(os.Args) < 2 {
		fmt.Println("File Not provided")
		return
	}
	filepath := os.Args[1]
	f, err := excelize.OpenFile(filepath)
	if err != nil {
		fmt.Println(err)
		return
	}

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	rows, err := f.GetRows("CSF111_202425_01_GradeBook")

	if err != nil {
		fmt.Println(err)
	}

	cols, err := f.GetCols("CSF111_202425_01_GradeBook")
	if err != nil {
		fmt.Println(err)
	}

	fmt.Println("Validating Data....")
	for rowidx, row := range rows[1:] {
		var sum float64 = 0
		colidx := 4
		for _, cellValue := range row[4:] {
			if colidx == 8 {
				colidx++
				continue
			}
			if colidx == 10 {
				break
			}
			val, err := strconv.ParseFloat(cellValue, 64)
			if err != nil {
				continue
			}
			sum += val
			colidx++
		}

		cellstring := fmt.Sprintf("K%d", rowidx+2)
		expectedSumStr, err := f.GetCellValue("CSF111_202425_01_GradeBook", cellstring)

		if err != nil {
			fmt.Println("error getting sum value for cell ", cellstring)
		}
		expectedSum, err := strconv.ParseFloat(expectedSumStr, 64)

		if err != nil {
			fmt.Println("error getting sum value for cell ", cellstring)
		}
		if expectedSum != sum {
			fmt.Println("Wrong Total score for cell ", cellstring, ". The total score should be ", expectedSum)
		}
	}
	fmt.Println()

	fmt.Println("Calculating General Averages...")

	for colidx, col := range cols[4:] {
		var sum float64 = 0
		rowidx := 0
		for _, cellvalue := range col {
			if rowidx == 0 {
				fmt.Print(cellvalue, ":\t")
				rowidx++
				continue
			}
			if cellvalue == "" {
				break
			}
			val, err := strconv.ParseFloat(cellvalue, 64)
			if err != nil {
				fmt.Println("Error getting cell value for row, ", rowidx, " and column ", colidx)
			}
			sum += val
			rowidx++
		}
		fmt.Print(sum / float64(rowidx))
		fmt.Println()
	}

	fmt.Println()

	MyBatchSlice := get24batchSlice(f)

	fmt.Println("Calculating Branch Wise Averages...")

	for colidx, col := range cols[4:] {
		var sum float64 = 0
		rowidx := 0
		for _, cellvalue := range col {
			if rowidx == 0 {
				fmt.Print(cellvalue, ":\t")
				rowidx++
				continue
			}
			if cellvalue == "" {
				break
			}
			for i := 0; i < len(MyBatchSlice); i++ {
				if rowidx == MyBatchSlice[i] {
					val, err := strconv.ParseFloat(cellvalue, 64)
					if err != nil {
						fmt.Println("Error getting cell value for row, ", rowidx, " and column ", colidx)
					}
					sum += val
					break
				}
			}
			rowidx++
		}
		fmt.Print(sum / float64(rowidx))
		fmt.Println()
	}
	fmt.Println()

	fmt.Println("Top 3 Students for each component: ")

	myvar := "E"
	for _, col := range cols[4:] {
		cellString := fmt.Sprintf("%s1", myvar)
		firstcolvalue, err := f.GetCellValue("CSF111_202425_01_GradeBook", cellString)
		if err != nil {
			fmt.Println(err)
		}
		fmt.Print("The Top 3 students for ", firstcolvalue, " are:\n")

		type StudentScore struct {
			rowIdx int
			score  float64
		}

		studentScores := make([]StudentScore, 0)

		for i := 1; i < len(col); i++ {
			score, err := strconv.ParseFloat(col[i], 64)
			if err != nil {
				continue
			}
			studentScores = append(studentScores, StudentScore{
				rowIdx: i,
				score:  score,
			})
		}

		for i := 0; i < len(studentScores); i++ {
			for j := i + 1; j < len(studentScores); j++ {
				if studentScores[i].score < studentScores[j].score {
					studentScores[i], studentScores[j] = studentScores[j], studentScores[i]
				}
			}
		}

		for i := 0; i < 3; i++ {
			empIDCol := "C"
			empIDRow := studentScores[i].rowIdx + 1
			empIDCell := fmt.Sprintf("%s%d", empIDCol, empIDRow)

			empID, err := f.GetCellValue("CSF111_202425_01_GradeBook", empIDCell)
			if err != nil {
				fmt.Println("Error getting EMPID:", err)
				continue
			}

			fmt.Printf("%d. EMPID: %s, Score: %.2f\n", i+1, empID, studentScores[i].score)
		}

		fmt.Println()
		myvar = string(myvar[0] + 1)
	}
}
