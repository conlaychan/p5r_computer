package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"sort"
	"strconv"
)

const materialFile = "Persona5Royal_material.xlsx"
const resultFile = "Persona5Royal_computed.xlsx"

var personas = make([]Persona, 0)
var personaByName = make(map[string]Persona)
var personaByArcana = make(map[string][]Persona)
var preciousIncr = make(map[string]map[string]int)
var materialArcanaMap = make(map[MaterialPair]string)
var materialSpecialPersonaMap = make(map[MaterialPair]string)
var resultMap = make(map[Persona]map[Persona]Persona)

func Persona5RoyalComputer() {
	fmt.Println("Persona 5 Royal：")
	if _, err := os.Lstat(resultFile); err == nil {
		if err := os.Remove(resultFile); err != nil {
			fmt.Println(err)
			fmt.Println("请先删除当前目录下的文件：" + resultFile)
			return
		}
	}
	fmt.Println("正在读取和分析素材文件：" + materialFile)
	readPersonas()
	fmt.Println("正在计算合成结果。。。")
	compute()
	fmt.Println("正在保存计算结果：" + resultFile)
	initColumns()
	saveExcel()
	fmt.Println("完成！")
}

func saveExcel() {
	sheet := "合成表"
	file := excelize.NewFile()
	defer func() {
		if err := file.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	if _, err := file.NewSheet(sheet); err != nil {
		fmt.Println(err)
		os.Exit(-1)
	}
	if err := file.DeleteSheet("sheet1"); err != nil {
		fmt.Println(err)
		os.Exit(-1)
	}

	for i, p := range personas {
		setCellValueOrExit(file, sheet, columns[i+2]+"1", p.arcana+"Lv"+strconv.FormatInt(p.level, 10))
		setCellValueOrExit(file, sheet, columns[i+2]+"2", p.name)
	}
	for rowNum, p1 := range personas {
		for columnIndex, p2 := range personas {
			if columnIndex == 0 {
				setCellValueOrExit(file, sheet, columns[columnIndex]+strconv.Itoa(rowNum+3), p1.arcana+"Lv"+strconv.FormatInt(p1.level, 10))
				setCellValueOrExit(file, sheet, columns[columnIndex+1]+strconv.Itoa(rowNum+3), p1.name)
			}
			result, exist := resultMap[p1][p2]
			if exist {
				setCellValueOrExit(file, sheet, columns[columnIndex+2]+strconv.Itoa(rowNum+3), result.arcana+"Lv"+strconv.FormatInt(result.level, 10)+" "+result.name)
			}
		}
	}

	if err := file.SaveAs(resultFile); err != nil {
		fmt.Println(err)
		os.Exit(-1)
	}
}

func setCellValueOrExit(file *excelize.File, sheet string, cell string, value string) {
	err := file.SetCellValue(sheet, cell, value)
	if err != nil {
		fmt.Println(err)
		os.Exit(-1)
	}
}

var words = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
var columns = make([]string, 0)

func initColumns() {
	for _, letter1 := range words {
		columns = append(columns, string(letter1))
	}
	for _, letter1 := range words {
		for _, letter2 := range words {
			columns = append(columns, string(letter1)+string(letter2))
		}
	}
}

func compute() {
	for _, p1 := range personas {
		arcana1 := p1.arcana
		name1 := p1.name
		for _, p2 := range personas {
			arcana2 := p2.arcana
			name2 := p2.name
			result := new(Persona)
			if name1 == name2 {
				result = nil
			} else if p1.isPrecious && p2.isPrecious {
				// 两个宝魔，不可合成
				result = nil
			} else if p1.isPrecious || p2.isPrecious {
				// 有且只有一个宝魔
				var preciousName string // 宝魔名称
				var arcana string       // 战斗面具的塔罗牌
				var materialPersona string
				if p1.isPrecious {
					preciousName = name1
					arcana = arcana2
					materialPersona = name2
				} else {
					preciousName = name2
					arcana = arcana1
					materialPersona = name1
				}
				preciousMap := preciousIncr[arcana]
				if preciousMap == nil || preciousMap[preciousName] == 0 {
					result = nil
				} else {
					incr := preciousMap[preciousName]
					results := personaByArcana[arcana]
					materialIndex := -1
					for i, persona := range results {
						if persona.name == materialPersona {
							materialIndex = i
							break
						}
					}
					targetIndex := materialIndex + incr
					result = nil
					for targetIndex >= 0 && targetIndex < len(results) {
						target := results[targetIndex]
						if isNotSpecial(target, name1, name2) {
							result = &target
							break
						} else {
							var shit int
							if incr > 0 {
								shit = 1
							} else {
								shit = -1
							}
							targetIndex = targetIndex + shit
						}
					}
				}
			} else {
				// 没有宝魔
				// 检查特殊合成
				checkSpecial := checkSpecial(name1, name2)
				if checkSpecial != nil {
					result = checkSpecial
				} else {
					// 标准合成
					resultArcana := materialArcanaMap[MaterialPair{arcana1, arcana2}]
					if resultArcana == "" {
						result = nil
					} else {
						level := (p1.level+p2.level)/2 + 1
						results := personaByArcana[resultArcana]
						if arcana1 == arcana2 {
							// 同种塔罗牌，往低段位找
							result = findLower(results, level, name1, name2)
						} else {
							// 不同种塔罗牌，往高段位找
							result = findHigher(results, level, name1, name2)
							if result == nil {
								// 往高段位找不到，就改为往低段位找
								result = findLower(results, level, name1, name2)
							}
						}
					}
				}
			}
			if resultMap[p1] == nil {
				resultMap[p1] = make(map[Persona]Persona)
			}
			if result != nil {
				resultMap[p1][p2] = *result
			}
		}
	}
}

func findHigher(results []Persona, level int64, name1 string, name2 string) *Persona {
	for _, persona := range results {
		if persona.level >= level && isNotSpecial(persona, name1, name2) {
			return &persona
		}
	}
	return nil
}

func findLower(results []Persona, level int64, name1 string, name2 string) *Persona {
	for i := len(results) - 1; i >= 0; i-- {
		persona := results[i]
		if persona.level <= level && isNotSpecial(persona, name1, name2) {
			return &persona
		}
	}
	return nil
}

func checkSpecial(name1 string, name2 string) *Persona {
	result := materialSpecialPersonaMap[MaterialPair{name1, name2}]
	if result == "" {
		result = materialSpecialPersonaMap[MaterialPair{name2, name1}]
	}
	if result == "" {
		return nil
	}
	p, exist := personaByName[result]
	if exist {
		return &p
	} else {
		return nil
	}
}

func isNotSpecial(persona Persona, name1 string, name2 string) bool {
	name := persona.name
	return !containsValue(materialSpecialPersonaMap, name) && !persona.isGroup && name != name1 && name != name2 && !persona.isPrecious
}

func containsValue[M ~map[K]V, K, V comparable](m M, v V) bool {
	for _, x := range m {
		if x == v {
			return true
		}
	}
	return false
}

func readPersonas() {
	file, err := excelize.OpenFile(materialFile)
	if err != nil {
		fmt.Println("无法读取素材文件：" + materialFile)
		fmt.Println(err)
		os.Exit(-1)
	}
	defer func() {
		if err := file.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	sheetList := file.GetSheetList()
	rows, err := file.GetRows(sheetList[0])
	if err != nil {
		fmt.Println("无法读取素材文件：" + materialFile + "，工作表：" + sheetList[0])
		fmt.Println(err)
		os.Exit(-1)
	}
	for rowNum, row := range rows {
		if rowNum == 0 {
			continue
		}
		persona := Persona{
			arcana:     row[0],
			name:       row[2],
			isPrecious: row[3] == "宝魔",
			isGroup:    row[3] == "集体断头台",
		}
		level, _ := strconv.ParseInt(row[1], 10, 64)
		persona.level = level
		personas = append(personas, persona)
	}
	// 宝魔升降
	rows, err = file.GetRows(sheetList[1])
	if err != nil {
		fmt.Println("无法读取素材文件：" + materialFile + "，工作表：" + sheetList[1])
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
				if preciousIncr[arcana] == nil {
					preciousIncr[arcana] = make(map[string]int)
				}
				preciousIncr[arcana][preciousName] = int(incr)
			}
		}
	}
	// 塔罗牌二合一
	rows, err = file.GetRows(sheetList[2])
	if err != nil {
		fmt.Println("无法读取素材文件：" + materialFile + "，工作表：" + sheetList[2])
		fmt.Println(err)
		os.Exit(-1)
	}
	rowMap := make(map[int]string)
	for rowNum, row := range rows {
		arcana1 := ""
		for columnIndex, cell := range row {
			if rowNum == 0 {
				if columnIndex != 0 {
					rowMap[columnIndex] = cell
				}
			} else if columnIndex == 0 {
				arcana1 = cell
			} else {
				arcana2 := rowMap[columnIndex]
				materialArcanaMap[MaterialPair{arcana1, arcana2}] = cell
			}
		}
	}
	// 特殊二体合成
	rows, err = file.GetRows(sheetList[3])
	if err != nil {
		fmt.Println("无法读取素材文件：" + materialFile + "，工作表：" + sheetList[3])
		fmt.Println(err)
		os.Exit(-1)
	}
	for rowNum, row := range rows {
		if rowNum != 0 {
			m1 := row[0]
			m2 := row[1]
			result := row[2]
			materialSpecialPersonaMap[MaterialPair{m1, m2}] = result
		}
	}
	sort.Sort(PersonaSlice(personas))
	for _, persona := range personas {
		personaByName[persona.name] = persona
		if personaByArcana[persona.arcana] == nil {
			personaByArcana[persona.arcana] = make([]Persona, 0)
		}
		personaByArcana[persona.arcana] = append(personaByArcana[persona.arcana], persona)
	}
}

type Persona struct {
	arcana     string
	level      int64
	name       string
	isPrecious bool
	isGroup    bool
}

type PersonaSlice []Persona

func (p PersonaSlice) Len() int {
	return len(p)
}

func (p PersonaSlice) Less(i, j int) bool {
	if p[i].level != p[j].level {
		return p[i].level < p[j].level
	} else {
		return p[i].name < p[j].name
	}
}

func (p PersonaSlice) Swap(i, j int) {
	p[i], p[j] = p[j], p[i]
}

type MaterialPair struct {
	m1 string
	m2 string
}
