package main

import (
	"fmt"
	"log"
	"os"
	"sort"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

var (
	total     = 100
	inputFile = "order.xlsx"
)

func init() {
	args := os.Args[1:]
	switch {
	case len(args) == 1:
		if strings.HasSuffix(args[0], "xlsx") {
			inputFile = args[0]
		} else {
			t, err := strconv.Atoi(args[0])
			if err != nil {
				log.Println("please input a valid number, use default: ", total)
				return
			}
			total = t
		}
	case len(args) > 1:
		if strings.HasSuffix(args[0], "xlsx") {
			inputFile = args[0]
		}
		t, err := strconv.Atoi(args[1])
		if err != nil {
			log.Println("please input a valid number, use default: ", total)
			return
		}
		total = t
	}
}

func main() {
	log.Println("1. read file: ", inputFile)
	f, err := excelize.OpenFile(inputFile)
	if err != nil {
		fmt.Println(err)
		return
	}
	// Get all the rows in the Sheet1.
	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}
	log.Println("2. read map")
	skuIdMap := make(map[string]int)
	for _, row := range rows {
		skuIdMap[row[0]] = 0
	}

	log.Println("3. read sku list")
	// 获取所有 sku 的列表和他们对应的位置
	// 这里的顺序是不稳定的，如果需要稳定的话这里需要先排序
	var skuIds []string
	var skuOrder = 0
	for skuId := range skuIdMap {
		skuIds = append(skuIds, skuId)
		skuIdMap[skuId] = skuOrder
		skuOrder++
	}

	log.Println("4. init matrix")
	// 初始化矩阵
	matrix := [][]int{}
	for i := 0; i < len(skuIds); i++ {
		matrix = append(matrix, make([]int, len(skuIds)))
	}

	// 遍历每一行数据
	// 这里要求订单号必须是排好序的

	log.Println("5. iter order", len(rows), len(skuIds))
	for start := 0; start < len(rows); {
		i := 1
		for start+i < len(rows) && rows[start][1] == rows[start+i][1] {
			i++
		}
		// 如果一个订单只有一个 sku，直接跳到下一个
		if i == 1 {
			start++
			continue
		}
		var skuInOrder = []string{}
		for j := start; j < start+i; j++ {
			skuInOrder = append(skuInOrder, rows[j][0])
		}
		for _, sku1 := range skuInOrder {
			for _, sku2 := range skuInOrder {
				matrix[skuIdMap[sku1]][skuIdMap[sku2]] += 1
			}
		}
		start += i
	}

	log.Println("6. construct sku pairs")
	var pairs []SkuPair
	for i := 0; i < len(matrix); i++ {
		for j := i + 1; j < len(matrix); j++ {
			pairs = append(pairs, SkuPair{
				Sku1:  skuIds[i],
				Sku2:  skuIds[j],
				Count: matrix[i][j],
			})
		}
	}

	log.Println("7. sort pairs", len(pairs))
	sort.SliceStable(pairs, func(i, j int) bool {
		return pairs[i].Count > pairs[j].Count
	})

	log.Println("8. save first", total, "lines")
	save(pairs)

	log.Println("finished")
}

func save(pairs []SkuPair) {
	f := excelize.NewFile()
	// Create a new sheet.
	index := f.NewSheet("Sheet1")
	// Set value of a cell.
	for i := 0; i < total; i++ {
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i+1), pairs[i].Sku1)
		f.SetCellValue("Sheet1", fmt.Sprintf("B%d", i+1), pairs[i].Sku2)
		f.SetCellValue("Sheet1", fmt.Sprintf("C%d", i+1), pairs[i].Count)
	}
	// Set active sheet of the workbook.
	f.SetActiveSheet(index)
	// Save spreadsheet by the given path.
	if err := f.SaveAs(fmt.Sprintf("Sku-%s.xlsx", time.Now().Format("2006-01-02 15-04-05"))); err != nil {
		fmt.Println(err)
	}
}

type SkuPair struct {
	Sku1  string
	Sku2  string
	Count int
}
