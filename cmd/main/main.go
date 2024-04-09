package main

import (
	"encoding/hex"
	"encoding/xml"
	"flag"
	"fmt"
	"log"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strconv"
	"strings"

	"github.com/antchfx/xmlquery"
	"github.com/atotto/clipboard"
	"github.com/xuri/excelize/v2"
)

var (
	xlsxFile *excelize.File
)

type fmxmlSnippet struct {
	XMLName   xml.Name `xml:"fmxmlsnippet"`
	Type      string   `xml:"type,attr"`
	BaseTable struct {
		Name  string `xml:"name,attr"`
		Field struct {
			ID          string `xml:"id,attr"`
			DataType    string `xml:"dataType,attr"`
			FieldType   string `xml:"fieldType,attr"`
			Name        string `xml:"name,attr"`
			Calculation struct {
				XMLName xml.Name `xml:"Calculation"`
				Value   string   `xml:",cdata"`
			}
			Comment   string `xml:"Comment"`
			AutoEnter struct {
				OverwriteExistingValue string `xml:"overwriteExistingValue,attr"`
				AlwaysEvaluate         string `xml:"alwaysEvaluate,attr"`
				AllowEditing           string `xml:"allowEditing,attr"`
				Constant               string `xml:"constant,attr"`
				Furigana               string `xml:"furigana,attr"`
				Lookup                 string `xml:"lookup,attr"`
				Calculation            string `xml:"calculation,attr"`
				ConstantData           string `xml:"ConstantData"`
			} `xml:"AutoEnter"`
			Validation struct {
				Message                   string `xml:"message,attr"`
				MaxLength                 string `xml:"maxLength,attr"`
				Valuelist                 string `xml:"valuelist,attr"`
				Calculation               string `xml:"calculation,attr"`
				AlwaysValidateCalculation string `xml:"alwaysValidateCalculation,attr"`
				Type                      string `xml:"type,attr"`
				NotEmpty                  struct {
					Value string `xml:"value,attr"`
				} `xml:"NotEmpty"`
				Unique struct {
					Value string `xml:"value,attr"`
				} `xml:"Unique"`
				Existing struct {
					Value string `xml:"value,attr"`
				} `xml:"Existing"`
				MaxDataLength struct {
					Value string `xml:"value,attr"`
				} `xml:"MaxDataLength"`
				StrictDataType struct {
					Value string `xml:"value,attr"`
				} `xml:"StrictDataType"`
				StrictValidation struct {
					Value string `xml:"value,attr"`
				} `xml:"StrictValidation"`
			} `xml:"Validation"`
			Storage struct {
				AutoIndex     string `xml:"autoIndex,attr"`
				Index         string `xml:"index,attr"`
				IndexLanguage string `xml:"indexLanguage,attr"`
				Global        string `xml:"global,attr"`
				MaxRepetition string `xml:"maxRepetition,attr"`
			} `xml:"Storage"`
		} `xml:"Field"`
	} `xml:"BaseTable"`
}

func returnCellValue(sheetName string, rowAxis int, cellName string, defaultValue string) string {
	var cellValue string
	if len(cellName) > 0 {
		colAxis, _, _ := excelize.CellNameToCoordinates(cellName)
		cellLabel, _ := excelize.CoordinatesToCellName(colAxis, rowAxis+1)
		cellValue, _ = xlsxFile.GetCellValue(sheetName, cellLabel)
		//cellValue = row[colAxis-1]
	}
	if len(cellValue) == 0 {
		cellValue = defaultValue
	}
	// --------------------------------------------------
	cellValue = strings.Replace(cellValue, "通常", "Normal", -1)
	cellValue = strings.Replace(cellValue, "計算", "Calculated", -1)
	cellValue = strings.Replace(cellValue, "集計", "Summary", -1)
	cellValue = strings.Replace(cellValue, "テキスト", "Text", -1)
	cellValue = strings.Replace(cellValue, "数字", "Number", -1)
	cellValue = strings.Replace(cellValue, "日付", "Date", -1)
	cellValue = strings.Replace(cellValue, "時刻", "Time", -1)
	cellValue = strings.Replace(cellValue, "タイムスタンプ", "TimeStamp", -1)
	cellValue = strings.Replace(cellValue, "オブジェクト", "Binary", -1)
	cellValue = strings.Replace(cellValue, "数字のみ", "Numeric", -1)
	cellValue = strings.Replace(cellValue, "日付のみ", "FourDigitYear", -1)
	cellValue = strings.Replace(cellValue, "時刻のみ", "TimeOfDay", -1)
	return cellValue
}

func main() {
	// fmlXMLSnippet を定義
	var rec fmxmlSnippet

	// 実行ファイルのパスを取得
	exe, err := os.Executable()
	if err != nil {
		log.Fatal(err)
	}
	dir := filepath.Dir(exe)

	//ログファイルを作成
	file, err := os.OpenFile(dir+"/debug.log", os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
	if err != nil {
		log.Fatal(err)
	}
	defer file.Close()
	log.SetOutput(file)

	// config ファイルの読み込み
	r, err := os.Open(dir + "/config.xml")
	if err != nil {
		log.Println(err)
	}
	defer r.Close()

	// https: //zenn.dev/nnabeyang/scraps/a1429f7f1214e9
	dec := xml.NewDecoder(r)
	err = dec.Decode(&rec)
	if err != nil {
		log.Println(err)
	}

	// xlsx の読み込み
	flag.Parse()
	xlsxFile, err = excelize.OpenFile(flag.Arg(0))
	if err != nil {
		log.Println(err)
	}
	defer xlsxFile.Close()
	// root エレメントの作成
	rootElement := &xmlquery.Node{
		Data: "fmxmlsnippet",
		Type: xmlquery.ElementNode,
		Attr: []xmlquery.Attr{
			{Name: xml.Name{Local: "type"}, Value: "FMObjectList"},
		},
	}

	for index, sheetName := range xlsxFile.GetSheetMap() {
		// シート名が #SAMPLE の場合は処理をスキップ
		fmt.Println(index, sheetName)
		if sheetName == "#SAMPLE" {
			continue
		}
		// シート内容が空の場合は処理をスキップ
		rows, _ := xlsxFile.GetRows(sheetName)
		if len(rows) == 0 {
			continue
		}
		// フィールド内容座標が取得できない場合は処理をスキップ
		_, rowAxis, err := excelize.SplitCellName(rec.BaseTable.Field.ID)
		if err != nil {
			log.Println(rowAxis, err)
			continue
		}
		// table エレメントの作成
		cellValue, _ := xlsxFile.GetCellValue(sheetName, rec.BaseTable.Name)
		tableElement := &xmlquery.Node{
			Data: "BaseTable",
			Type: xmlquery.ElementNode,
			Attr: []xmlquery.Attr{
				{Name: xml.Name{Local: "name"}, Value: cellValue},
			},
		}
		// field エレメントの作成
		fieldXML := rec.BaseTable.Field
		// 各行の末尾にある継続的な空白のセルはスキップされるため、各行の長さが不一致になる可能性がある、シートからセルの値を直接取得する
		for rowIndex, row := range rows {
			if rowIndex >= (rowAxis-1) && len(row) > 1 {
				fmt.Println(rowIndex, row)
				// --------------------------------------------------
				fieldElement := &xmlquery.Node{
					Data: "Field",
					Type: xmlquery.ElementNode,
					Attr: []xmlquery.Attr{
						{Name: xml.Name{Local: "id"}, Value: returnCellValue(sheetName, rowIndex, fieldXML.ID, strconv.Itoa(rowIndex))},
						{Name: xml.Name{Local: "name"}, Value: returnCellValue(sheetName, rowIndex, fieldXML.Name, fmt.Sprintf("Field#%d", rowIndex))},
						{Name: xml.Name{Local: "fieldType"}, Value: returnCellValue(sheetName, rowIndex, fieldXML.FieldType, "Normal")},
						{Name: xml.Name{Local: "dataType"}, Value: returnCellValue(sheetName, rowIndex, fieldXML.DataType, "Text")},
					},
				}
				// --------------------------------------------------
				commentElement := &xmlquery.Node{
					Data: "Comment",
					Type: xmlquery.ElementNode,
				}
				xmlquery.AddChild(commentElement, &xmlquery.Node{
					Data: returnCellValue(sheetName, rowIndex, fieldXML.Comment, ""),
					Type: xmlquery.TextNode,
				})
				xmlquery.AddChild(fieldElement, commentElement)
				// --------------------------------------------------
				autoEnterElement := &xmlquery.Node{
					Data: "AutoEnter",
					Type: xmlquery.ElementNode,
					Attr: []xmlquery.Attr{
						/* True, False*/
						{Name: xml.Name{Local: "constant"}, Value: "False"},
						{Name: xml.Name{Local: "calculation"}, Value: "False"},
						{Name: xml.Name{Local: "alwaysEvaluate"}, Value: returnCellValue(sheetName, rowIndex, fieldXML.AutoEnter.AlwaysEvaluate, "False")},
						{Name: xml.Name{Local: "overwriteExistingValue"}, Value: returnCellValue(sheetName, rowIndex, fieldXML.AutoEnter.OverwriteExistingValue, "True")},
						{Name: xml.Name{Local: "allowEditing"}, Value: returnCellValue(sheetName, rowIndex, fieldXML.AutoEnter.AllowEditing, "True")},
						{Name: xml.Name{Local: "furigana"}, Value: returnCellValue(sheetName, rowIndex, fieldXML.AutoEnter.Furigana, "False")},
						{Name: xml.Name{Local: "lookup"}, Value: returnCellValue(sheetName, rowIndex, fieldXML.AutoEnter.Lookup, "False")},
					},
				}
				// --------------------------------------------------
				if returnCellValue(sheetName, rowIndex, fieldXML.AutoEnter.Constant, "固定値") == "固定値" {
					autoEnterElement.SetAttr("constant", "True")
					constantDataElement := &xmlquery.Node{
						Data: "ConstantData",
						Type: xmlquery.ElementNode,
					}
					xmlquery.AddChild(constantDataElement, &xmlquery.Node{
						Data: returnCellValue(sheetName, rowIndex, fieldXML.AutoEnter.ConstantData, ""),
						Type: xmlquery.TextNode,
					})
					xmlquery.AddChild(autoEnterElement, constantDataElement)
				} else {
					autoEnterElement.SetAttr("calculation", "True")
					calculationDataElement := &xmlquery.Node{
						Data: "Calculation",
						Type: xmlquery.ElementNode,
						Attr: []xmlquery.Attr{
							{Name: xml.Name{Local: "table"}, Value: ""},
						},
					}
					xmlquery.AddChild(calculationDataElement, &xmlquery.Node{
						Data: returnCellValue(sheetName, rowIndex, fieldXML.AutoEnter.ConstantData, ""),
						Type: xmlquery.CharDataNode,
					})
					xmlquery.AddChild(autoEnterElement, calculationDataElement)
				}
				xmlquery.AddChild(fieldElement, autoEnterElement)
				// --------------------------------------------------
				validationElement := &xmlquery.Node{
					Data: "Validation",
					Type: xmlquery.ElementNode,
					Attr: []xmlquery.Attr{
						/* True, False*/
						{Name: xml.Name{Local: "maxLength"}, Value: "False"},
						{Name: xml.Name{Local: "message"}, Value: returnCellValue(sheetName, rowIndex, fieldXML.Validation.Message, "False")},
						{Name: xml.Name{Local: "valuelist"}, Value: returnCellValue(sheetName, rowIndex, fieldXML.Validation.Valuelist, "False")},
						{Name: xml.Name{Local: "calculation"}, Value: returnCellValue(sheetName, rowIndex, fieldXML.Validation.Calculation, "False")},
						{Name: xml.Name{Local: "alwaysValidateCalculation"}, Value: returnCellValue(sheetName, rowIndex, fieldXML.Validation.AlwaysValidateCalculation, "False")},
						{Name: xml.Name{Local: "type"}, Value: returnCellValue(sheetName, rowIndex, "", "OnlyDuringDataEntry")},
					},
				}
				/* True, False */
				uniqueElement := &xmlquery.Node{
					Data: "Unique",
					Type: xmlquery.ElementNode,
					Attr: []xmlquery.Attr{
						{Name: xml.Name{Local: "value"}, Value: returnCellValue(sheetName, rowIndex, fieldXML.Validation.Unique.Value, "True")},
					},
				}
				xmlquery.AddChild(validationElement, uniqueElement)
				/* True, False */
				notEmptyElement := &xmlquery.Node{
					Data: "NotEmpty",
					Type: xmlquery.ElementNode,
					Attr: []xmlquery.Attr{
						{Name: xml.Name{Local: "value"}, Value: returnCellValue(sheetName, rowIndex, fieldXML.Validation.NotEmpty.Value, "True")},
					},
				}
				xmlquery.AddChild(validationElement, notEmptyElement)
				/* 0 - 999 */
				maxLengthElement := &xmlquery.Node{
					Data: "MaxDataLength",
					Type: xmlquery.ElementNode,
					Attr: []xmlquery.Attr{
						{Name: xml.Name{Local: "value"}, Value: returnCellValue(sheetName, rowIndex, fieldXML.Validation.MaxDataLength.Value, "")},
					},
				}
				if len(returnCellValue(sheetName, rowIndex, fieldXML.Validation.MaxDataLength.Value, "")) > 0 {
					validationElement.SetAttr("maxLength", "True")
				}
				xmlquery.AddChild(validationElement, maxLengthElement)
				xmlquery.AddChild(fieldElement, validationElement)
				// --------------------------------------------------
				storageEnterElement := &xmlquery.Node{
					Data: "Storage",
					Type: xmlquery.ElementNode,
					Attr: []xmlquery.Attr{
						{Name: xml.Name{Local: "autoIndex"}, Value: returnCellValue(sheetName, rowIndex, fieldXML.Storage.AutoIndex, "True")},
						{Name: xml.Name{Local: "index"}, Value: returnCellValue(sheetName, rowIndex, fieldXML.Storage.Index, "None")},
						{Name: xml.Name{Local: "indexLanguage"}, Value: returnCellValue(sheetName, rowIndex, fieldXML.Storage.IndexLanguage, "Japanese")},
						{Name: xml.Name{Local: "global"}, Value: returnCellValue(sheetName, rowIndex, fieldXML.Storage.Global, "False")},
						{Name: xml.Name{Local: "maxRepetition"}, Value: returnCellValue(sheetName, rowIndex, fieldXML.Storage.MaxRepetition, "1")},
					},
				}
				xmlquery.AddChild(fieldElement, storageEnterElement)
				// --------------------------------------------------
				xmlquery.AddChild(tableElement, fieldElement)
			}
		}

		xmlquery.AddChild(rootElement, tableElement)
	}

	// log.Println(rootElement.OutputXML(true))

	switch runtime.GOOS {
	case "darwin":
		hexData := hex.EncodeToString([]byte(rootElement.OutputXML(true)))
		hexData = "XMTB" + hexData
		hexData = "«data " + hexData + "»"
		// https: //stackoverflow.com/questions/45248144
		err := exec.Command("/usr/bin/osascript", "-e", fmt.Sprintf(`set the clipboard to %s`, hexData)).Run()
		if err != nil {
			log.Println(err)
		}
	case "windows":
		err = clipboard.WriteAll(rootElement.OutputXML(true))
		if err != nil {
			log.Println(err)
		}
	default:
		log.Println("ohter")
	}

}

// OOS=darwin GOARCH=amd64 go build -o ../../build/generateTables
// GOOS=darwin GOARCH=arm64 go build -o ../../build/generateTables
// GOOS=windows GOARCH=amd64 go build -o ../../build/generateTables.ext
