package main

import (
	"fmt"

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

	selectedCols := make([]int, 6)
	selectedCols.append(4, 5, 6, 7, 8, 9)

	for _, row := range rows {

	}
}
