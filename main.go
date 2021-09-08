package main

import (
	"fmt"
	"io/ioutil"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"time"

	"github.com/xuri/excelize/v2"
	"onns.xyz/parse-excel/model"
)

/*
@Time : 2021/8/30 21:43
@Author : onns
@File : /main.go
*/

func getWorkDir() (res string) {
	ex, err := os.Executable()
	if err != nil {
		panic(err)
	}
	res = filepath.Dir(ex)
	return
}

func parseExcel(fileName string) (res map[string]*model.Info, date string) {
	res = make(map[string]*model.Info)
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		fmt.Println(err)
		return
	}
	rows, err := f.GetRows("支部数据")
	for i, row := range rows {
		if i == 0 {
			continue
		}
		if i == 1 {
			date = row[0]
		}
		n, _ := strconv.ParseInt(row[2], 10, 64)
		res[row[1]] = &model.Info{
			Date:       row[0],
			Number:     n,
			Engagement: row[3],
		}
	}
	return
}

func parseAxis(m, n int) (axis string) {
	return fmt.Sprintf("%c%d", m+'A', n)
}

func generateExcel(totalInfo map[string]map[string]*model.Info, dateList []string, fileName string) {
	f := excelize.NewFile()
	n := 1
	for branch, branchInfo := range totalInfo {
		n += 1
		f.SetCellValue("Sheet1", parseAxis(1, n), branch)
		for i, date := range dateList {
			f.SetCellValue("Sheet1", parseAxis(2, n), branchInfo[date].Number)
			f.SetCellValue("Sheet1", parseAxis(i+3, n), branchInfo[date].Engagement)
		}
	}
	// 根据指定路径保存文件
	if err := f.SaveAs(fileName); err != nil {
		fmt.Println(err)
	}
}
func main() {
	workDir := getWorkDir()
	items, _ := ioutil.ReadDir(workDir)
	totalInfo := make(map[string]map[string]*model.Info)
	dateList := make([]string, 0)
	for _, item := range items {
		if !item.IsDir() && filepath.Ext(item.Name()) == ".xlsx" {
			dailyInfo, date := parseExcel(filepath.Join(workDir, item.Name()))
			if date == "" {
				continue
			}
			dateList = append(dateList, date)
			for branch, info := range dailyInfo {
				branchInfo, ok := totalInfo[branch]
				if !ok {
					branchInfo = make(map[string]*model.Info)
				}
				branchInfo[date] = info
				totalInfo[branch] = branchInfo
			}
		}
	}
	sort.Sort(sort.StringSlice(dateList))
	generateExcel(totalInfo, dateList, filepath.Join(workDir, "res", fmt.Sprintf("汇总_%s.xlsx", time.Now().Format("2006-01-02 150405"))))
}
