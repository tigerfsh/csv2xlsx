package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/tealeg/xlsx"
)

var xlsxPath = flag.String("o", "", "Path to the XLSX output file")
var csvPath = flag.String("f", "", "Path to the CSV input file")
var delimiter = flag.String("d", ";", "Delimiter for felds in the CSV input.")

func usage() {
	fmt.Printf(`%s: -f=<CSV Input File> -o=<XLSX Output File> -d=<Delimiter>

`,
		os.Args[0])
}

func generateXLSXFromCSV(csvPath string, XLSXPath string, delimiter string) error {
	csvFile, err := os.Open(csvPath)
	if err != nil {
		return err
	}
	defer csvFile.Close()
	reader := csv.NewReader(csvFile)
	if len(delimiter) > 0 {
		reader.Comma = rune(delimiter[0])
	} else {
		reader.Comma = rune(';')
	}
	xlsxFile := xlsx.NewFile()
	sheet, err := xlsxFile.AddSheet(csvPath)
	if err != nil {
		return err
	}
	fields, err := reader.Read()
	for err == nil {
		lastElem := fields[len(fields)-1]

		row := sheet.AddRow()
		for _, field := range fields[:len(fields)-1] {
			cell := row.AddCell()
			cell.Value = field
		}
		if len(lastElem) != 0 && lastElem != "MergeCells" {

			fmt.Printf("lastElem: %v\n", lastElem)
			mergeRangeSlice := strings.Split(lastElem, ";")

			fmt.Printf("mergeRangeSlice: %v\n", mergeRangeSlice)

			for _, mergeRange := range mergeRangeSlice {
				fmt.Printf("mergeRange: %v\n", mergeRange)
				indexSlice := strings.Split(mergeRange, ",")
				fmt.Printf("indexSlice, len:%v, value:%v\n", len(indexSlice), indexSlice)
				fmt.Println(len(indexSlice))
				// fmt.Println(indexSlice)
				if len(indexSlice) == 4 {
					// merge cells
					indexIntSlice := []int{}
					for _, v := range indexSlice {
						tmp, _ := strconv.Atoi(v)
						indexIntSlice = append(indexIntSlice, tmp)
					}
					startH, startV, endH, endV := indexIntSlice[0], indexIntSlice[1], indexIntSlice[2], indexIntSlice[3]
					fmt.Println(startH, startV, endH, endV)
					for v := startV; v <= endV; v++ {
						cell := sheet.Cell(startH, v)
						cell.Merge(0, endH-startH)
					}
				}

			}
		}
		fields, err = reader.Read()
	}
	if err != nil {
		fmt.Printf(err.Error())
	}
	return xlsxFile.Save(XLSXPath)
}

func main() {
	flag.Parse()
	if len(os.Args) < 3 {
		usage()
		return
	}
	flag.Parse()
	err := generateXLSXFromCSV(*csvPath, *xlsxPath, *delimiter)
	if err != nil {
		fmt.Printf(err.Error())
		return
	}
}
