package persona

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"sort"
	"strconv"
)

func readMaterialPersonas(context *Context, file *excelize.File, sheetName string) {
	rows, err := file.GetRows(sheetName)
	if err != nil {
		fmt.Println("无法读取素材文件：" + context.materialFile + "，工作表：" + sheetName)
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
			isGroup:    row[3] == "集体合成",
		}
		level, _ := strconv.ParseInt(row[1], 10, 64)
		persona.level = level
		context.personas = append(context.personas, persona)
	}
}

func readArcanaRules(file *excelize.File, sheetName string, context *Context) {
	rows, err := file.GetRows(sheetName)
	if err != nil {
		fmt.Println("无法读取素材文件：" + context.materialFile + "，工作表：" + sheetName)
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
				context.materialArcanaMap[MaterialPair{arcana1, arcana2}] = cell
			}
		}
	}
}

func readSpecialRules(file *excelize.File, sheetName string, context *Context) {
	rows, err := file.GetRows(sheetName)
	if err != nil {
		fmt.Println("无法读取素材文件：" + context.materialFile + "，工作表：" + sheetName)
		fmt.Println(err)
		os.Exit(-1)
	}
	for rowNum, row := range rows {
		if rowNum != 0 {
			m1 := row[0]
			m2 := row[1]
			result := row[2]
			context.materialSpecialPersonaMap[MaterialPair{m1, m2}] = result
		}
	}
}

func findHigher(results []Persona, level int64, name1 string, name2 string, context *Context) *Persona {
	for _, persona := range results {
		if persona.level >= level && isNotSpecial(persona, name1, name2, context) {
			return &persona
		}
	}
	return nil
}

func findLower(results []Persona, level int64, name1 string, name2 string, context *Context) *Persona {
	for i := len(results) - 1; i >= 0; i-- {
		persona := results[i]
		if persona.level <= level && isNotSpecial(persona, name1, name2, context) {
			return &persona
		}
	}
	return nil
}

func isNotSpecial(persona Persona, name1 string, name2 string, context *Context) bool {
	name := persona.name
	return !MapContainsValue(context.materialSpecialPersonaMap, name) && !persona.isGroup && name != name1 && name != name2 && !persona.isPrecious
}

// 检查特殊二体合成
func checkSpecial(name1 string, name2 string, context *Context) *Persona {
	result := context.materialSpecialPersonaMap[MaterialPair{name1, name2}]
	if result == "" {
		result = context.materialSpecialPersonaMap[MaterialPair{name2, name1}]
	}
	if result == "" {
		return nil
	}
	p, exist := context.personaByName[result]
	if exist {
		return &p
	} else {
		return nil
	}
}

func sortMaterialPersonas(context *Context) {
	sort.Sort(PersonaSlice(context.personas))
	for _, persona := range context.personas {
		context.personaByName[persona.name] = persona
		if context.personaByArcana[persona.arcana] == nil {
			context.personaByArcana[persona.arcana] = make([]Persona, 0)
		}
		context.personaByArcana[persona.arcana] = append(context.personaByArcana[persona.arcana], persona)
	}
}

func MapContainsValue[M ~map[K]V, K, V comparable](m M, v V) bool {
	for _, x := range m {
		if x == v {
			return true
		}
	}
	return false
}

var words = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
var columns = make([]string, 0)

func InitColumns() {
	for _, letter1 := range words {
		columns = append(columns, string(letter1))
	}
	for _, letter1 := range words {
		for _, letter2 := range words {
			columns = append(columns, string(letter1)+string(letter2))
		}
	}
}

func setCellValueOrExit(file *excelize.File, sheet string, cell string, value string) {
	err := file.SetCellValue(sheet, cell, value)
	if err != nil {
		fmt.Println(err)
		os.Exit(-1)
	}
}

func saveExcel(context *Context) {
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

	for i, p := range context.personas {
		setCellValueOrExit(file, sheet, columns[i+2]+"1", p.arcana+"Lv"+strconv.FormatInt(p.level, 10))
		setCellValueOrExit(file, sheet, columns[i+2]+"2", p.name)
	}
	for rowNum, p1 := range context.personas {
		for columnIndex, p2 := range context.personas {
			if columnIndex == 0 {
				setCellValueOrExit(file, sheet, columns[columnIndex]+strconv.Itoa(rowNum+3), p1.arcana+"Lv"+strconv.FormatInt(p1.level, 10))
				setCellValueOrExit(file, sheet, columns[columnIndex+1]+strconv.Itoa(rowNum+3), p1.name)
			}
			result, exist := context.resultMap[p1][p2]
			if exist {
				setCellValueOrExit(file, sheet, columns[columnIndex+2]+strconv.Itoa(rowNum+3), result.arcana+"Lv"+strconv.FormatInt(result.level, 10)+" "+result.name)
			}
		}
	}

	if err := file.SaveAs(context.resultFile); err != nil {
		fmt.Println(err)
		os.Exit(-1)
	}
}

func compute(context *Context) {
	for _, p1 := range context.personas {
		arcana1 := p1.arcana
		name1 := p1.name
		for _, p2 := range context.personas {
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
				preciousMap := context.preciousIncr[arcana]
				if preciousMap == nil || preciousMap[preciousName] == 0 {
					result = nil
				} else {
					incr := preciousMap[preciousName]
					results := context.personaByArcana[arcana]
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
						if isNotSpecial(target, name1, name2, context) {
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
				specialResult := checkSpecial(name1, name2, context)
				if specialResult != nil {
					result = specialResult
				} else {
					// 标准合成
					resultArcana := context.materialArcanaMap[MaterialPair{arcana1, arcana2}]
					if resultArcana == "" {
						result = nil
					} else {
						level := (p1.level+p2.level)/2 + 1
						results := context.personaByArcana[resultArcana]
						if arcana1 == arcana2 {
							// 同种塔罗牌，往低段位找
							result = findLower(results, level, name1, name2, context)
						} else {
							// 不同种塔罗牌，往高段位找
							result = findHigher(results, level, name1, name2, context)
							if result == nil {
								// 往高段位找不到，就改为往低段位找
								result = findLower(results, level, name1, name2, context)
							}
						}
					}
				}
			}
			if context.resultMap[p1] == nil {
				context.resultMap[p1] = make(map[Persona]Persona)
			}
			if result != nil {
				context.resultMap[p1][p2] = *result
			}
		}
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

type Context struct {
	materialFile              string
	resultFile                string
	personas                  []Persona
	personaByName             map[string]Persona
	personaByArcana           map[string][]Persona
	preciousIncr              map[string]map[string]int
	materialArcanaMap         map[MaterialPair]string
	materialSpecialPersonaMap map[MaterialPair]string
	resultMap                 map[Persona]map[Persona]Persona
}
