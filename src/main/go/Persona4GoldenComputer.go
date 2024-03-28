package persona

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
)

func P4GoldenComputer() {
	fmt.Println("Persona 4 Golden：")

	var context = Context{
		materialFile:              "Persona4Golden_material.xlsx",
		resultFile:                "Persona4Golden_computed.xlsx",
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
	readPersonasP4G(&context)

	fmt.Println("正在计算合成结果。。。")
	compute(&context)

	fmt.Println("正在保存计算结果：" + context.resultFile)
	saveExcel(&context)

	fmt.Println("完成！")
}

func readPersonasP4G(context *Context) {
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

	// 塔罗牌二合一
	readArcanaRules(file, sheetList[1], context)

	// 特殊二体合成
	readSpecialRules(file, sheetList[2], context)

	sortMaterialPersonas(context)
}
