package main

import (
	"encoding/json"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"

	termbox "github.com/nsf/termbox-go"
)

func init() {
	if err := termbox.Init(); err != nil {
		panic(err)
	}
	termbox.SetCursor(0, 0)
	termbox.HideCursor()
}

type Data struct {
	Name           string      `json:"name"`
	Excel          interface{} `json: "excel"`
	AttributeNames []string    `json:"attributeNames"`
}

func main() {
	// files, err := ioutil.ReadDir("./table")
	// if err != nil {
	// 	panic("读取文件夹错误")
	// }

	var tableMap = make(map[string]interface{})

	err := filepath.Walk("./table",
		func(path string, f os.FileInfo, err error) error {
			if err != nil {
				return err
			}

			bo, err := regexp.MatchString(".xlsx", f.Name())

			if err != nil {
				panic("匹配出错")
			}

			if bo {
				fileExcel, err := excelize.OpenFile(path)
				if err != nil {
					panic("文件打开错误")
				}

				for _, sheetName := range fileExcel.GetSheetMap() {

					// 创建字典
					fileMap := make(map[string]interface{})

					tipsNames := []string{}
					attributeNames := []string{}
					attributeTypes := []string{}
					for i, row := range fileExcel.GetRows(sheetName) {
						if i == 0 {
							tipsNames = row
							continue
						}

						if i == 1 {
							attributeNames = row
							continue
						}

						if i == 2 {
							attributeTypes = row
							continue
						}

						sMap := make(map[string]interface{})
						aMap := []interface{}{}
						nMap := []string{}
						for j, s := range row {
							isTip, err := regexp.MatchString(`^#.*`, tipsNames[j])
							// println("是否是注释列：", isTip)
							if err != nil {
								panic("对比注释字段文字出错")
							}

							if isTip == true {
								continue
							}
							nMap = append(nMap, attributeNames[j])
							switch attributeTypes[j] {
							case "string":
								sMap[attributeNames[j]] = s
								aMap = append(aMap, s)
							case "int":
								v, err := strconv.Atoi(s)
								if err != nil {
									println("整型", sheetName, v, s, attributeNames[j])
									panic("字段解析错误")
								}
								sMap[attributeNames[j]] = v
								aMap = append(aMap, v)
							case "float":
								v, err := strconv.ParseFloat(s, 64)
								if err != nil {
									println("浮点型", sheetName, v)
									panic("字段解析错误")
								}

								sMap[attributeNames[j]] = v
								aMap = append(aMap, v)
							case "arraystring":
								v := strings.Split(s, "#")
								sMap[attributeNames[j]] = v
								aMap = append(aMap, v)
							case "arrayint":
								v := strings.Split(s, "#")
								intvs := []int{}
								for _, sid := range v {
									intv, err := strconv.Atoi(sid)
									if err == nil {
										intvs = append(intvs, intv)
									}
								}
								sMap[attributeNames[j]] = intvs
								aMap = append(aMap, intvs)
							case "arrayfloat":
								v := strings.Split(s, "#")
								intvs := []float64{}
								for _, sid := range v {
									intv, err := strconv.ParseFloat(sid, 64)
									if err == nil {
										intvs = append(intvs, intv)
									}
								}
								sMap[attributeNames[j]] = intvs
								aMap = append(aMap, intvs)
							}

						}
						fileMap[row[0]] = aMap
						// fileMap[row[0]] = sMap
						filePtr, err := os.Create("./assets/resources/config/" + sheetName + ".json")
						// fileBin, err := os.Create("./assets/resources/config/" + sheetName + ".bin")

						if err != nil {
							fmt.Println("文件创建失败", err.Error())
							// return
						}

						defer filePtr.Close()

						excelData := &Data{}
						excelData.Name = sheetName
						excelData.AttributeNames = nMap
						excelData.Excel = fileMap

						data, err := json.Marshal(excelData)
						// data, err := json.MarshalIndent(excelData, "", "\t")
						filePtr.Write(data)
					}

					println("成功导出: ", sheetName)

					tableMap[sheetName] = fileMap
				}

			}

			// fmt.Println(path, f.Size())
			return nil
		})
	if err != nil {
		log.Println(err)
	}

	// for _, f := range files {

	// }

	// filePtr, err := os.Create("./assets/resources/config/" + "table" + ".json")
	// fileBin, err := os.Create("./assets/resources/config/" + "table" + ".bin")
	// if err != nil {
	// 	fmt.Println("文件创建失败", err.Error())
	// 	return
	// }

	// defer fileBin.Close()

	// data, err := json.MarshalIndent(tableMap, "", "\t")
	// fileBin.Write(data)
	println("*****全部文件导出成功*****")

	pause()
}

func pause() {
	fmt.Println("请按任意键继续...")
Loop:
	for {
		switch ev := termbox.PollEvent(); ev.Type {
		case termbox.EventKey:
			break Loop
		}
	}
}
