package main

import (
	"fmt"
	"strconv"

	"github.com/xuri/excelize/v2"
)

func main() {
	f, err := excelize.OpenFile("CSF111_202425_01_GradeBook_stripped.xlsx")
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

}
