package main

import (
	"fmt"
	"io/ioutil"
	"os"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {

	xlsx := excelize.NewFile()
	//Check if the correct number of argument is supplied
	if len(os.Args) != 2 {
		fmt.Printf("shoudl be run as \n$%s testdir\n", os.Args[0])
		return
	}
	testDirs, err := ioutil.ReadDir(os.Args[1])
	if err != nil {
		fmt.Printf("Error reading the test directory name=%s err= %v", os.Args[1], err)
		return
	}
	ThisTest := RedisTest{Name: os.Args[1]}

	for _, dirName := range testDirs {
		if !dirName.IsDir() {
			fmt.Printf("Error: Not processing %s as its not a directory\n", dirName.Name())
			continue
		}
		fullDirName := os.Args[1] + "/" + dirName.Name()
		testFiles, err := ioutil.ReadDir(fullDirName)
		if err != nil {
			fmt.Printf("Error reading sub-directory name=%s err=%v\n", fullDirName, err)
			continue
		}
		fmt.Printf("Processing %s....\n", dirName.Name())
		sub := SubTest{Name: dirName.Name(), dufilename: fullDirName + "/" + "du.txt", topfilename: fullDirName + "/" + "top.txt"}
		for _, file := range testFiles {
			//fmt.Printf("processing %+v\n", file)
			if strings.Contains(file.Name(), "json") {
				jsFullFileName := fullDirName + "/" + file.Name()
				js, err := read_memtier_json(jsFullFileName)
				if err != nil {
					fmt.Printf("Unable to parse memtier output json filename =%s err=%v\n", jsFullFileName, err)
					continue
				}
				sub.memout = append(sub.memout, js)
			}
		}
		plotSheet(xlsx, sub.Name, sub)
		ThisTest.Sub = append(ThisTest.Sub, sub)
	}
	ThisTest.Summary(xlsx, "sheet1")
	err = xlsx.SaveAs(ThisTest.Name + ".xlsx")
	if err != nil {
		fmt.Printf("Error saving the file err = %v", err)
		return
	}
	fmt.Printf("Report generated\n")
}
