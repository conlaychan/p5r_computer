package persona

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"strconv"
)

func P5RoyalComputer() {
	fmt.Println("Persona 5 Royal：")

	var context = Context{
		materialFile:              "Persona5Royal_material.xlsx",
		resultFile:                "Persona5Royal_computed.xlsx",
		personas:                  make([]Persona, 0),
		personaByName:             make(map[string]Persona),
		personaByArcana:           make(map[string][]Persona),
		preciousIncr:              make(map[string]map[string]int),
		materialArcanaMap:         make(map[MaterialPair]string),
		materialSpecialPersonaMap: make(map[MaterialPair]string),
		resultMap:                 make(map[Persona]map[Persona]Persona),
	}
	if _, err := os.Lstat(context.resultFile); err == nil {
		if err := os.Remove(context.resultFile); err != nil {
			fmt.Println(err)
			fmt.Println("请先删除当前目录下的文件：" + context.resultFile)
			return
		}
	}

	fmt.Println("正在读取和分析素材文件：" + context.materialFile)
	readPersonasP5R(&context)

	fmt.Println("正在计算合成结果。。。")
	compute(&context)

	fmt.Println("正在保存计算结果：" + context.resultFile)
	saveExcel(&context)

	fmt.Println("完成！")
}

func readPersonasP5R(context *Context) {
	file, err := excelize.OpenFile(context.materialFile)
	if err != nil {
		fmt.Println("无法读取素材文件：" + context.materialFile)
		fmt.Println(err)
		os.Exit(-1)
	}
	defer func() {
		if err := file.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	sheetList := file.GetSheetList()
	readMaterialPersonas(context, file, sheetList[0])

	// 宝魔升降
	rows, err := file.GetRows(sheetList[1])
	if err != nil {
		fmt.Println("无法读取素材文件：" + context.materialFile + "，工作表：" + sheetList[1])
		fmt.Println(err)
		os.Exit(-1)
	}
	arcanaMap := make(map[int]string)
	for columnIndex, cell := range rows[0] {
		if columnIndex != 0 {
			arcanaMap[columnIndex] = cell
		}
	}
	for rowNum, row := range rows {
		if rowNum == 0 {
			continue
		}
		preciousName := "" // 宝魔面具
		for columnIndex, cell := range row {
			if columnIndex == 0 {
				preciousName = cell
			} else {
				incr, _ := strconv.ParseInt(cell, 10, 32)
				arcana := arcanaMap[columnIndex] // 战斗面具的塔罗牌
				if context.preciousIncr[arcana] == nil {
					context.preciousIncr[arcana] = make(map[string]int)
				}
				context.preciousIncr[arcana][preciousName] = int(incr)
			}
		}
	}
	// 塔罗牌二合一
	readArcanaRules(file, sheetList[2], context)

	// 特殊二体合成
	readSpecialRules(file, sheetList[3], context)

	sortMaterialPersonas(context)
}
