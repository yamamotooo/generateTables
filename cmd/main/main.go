package main

import (
	"bytes"
	"encoding/hex"
	"encoding/xml"
	"flag"
	"fmt"
	"io"
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

var cellValueReplacer = strings.NewReplacer(
	"通常タイプ", "Normal",
	"計算タイプ", "Calculated",
	"集計タイプ", "Summary",
	"テキスト型", "Text",
	"数字型", "Number",
	"日付型", "Date",
	"時刻型", "Time",
	"タイムスタンプ型", "TimeStamp",
	"オブジェクト型", "Binary",
	"数字のみ", "Numeric",
	"日付のみ", "FourDigitYear",
	"時刻のみ", "TimeOfDay",
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
				Table   string   `xml:"table,attr"`
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

func prettyXML(src string) string {
	var buf bytes.Buffer
	buf.WriteString(xml.Header)
	enc := xml.NewEncoder(&buf)
	enc.Indent("", "\t")
	dec := xml.NewDecoder(strings.NewReader(src))
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return src
		}
		if err = enc.EncodeToken(tok); err != nil {
			return src
		}
	}
	if err := enc.Flush(); err != nil {
		return src
	}
	return buf.String()
}

func returnCellValue(f *excelize.File, sheetName string, rowAxis int, cellName string, defaultValue string) string {
	var cellValue string
	if cellName != "" {
		colAxis, _, _ := excelize.CellNameToCoordinates(cellName)
		cellLabel, _ := excelize.CoordinatesToCellName(colAxis, rowAxis+1)
		cellValue, _ = f.GetCellValue(sheetName, cellLabel)
	}
	if cellValue == "" {
		cellValue = defaultValue
	}
	return cellValueReplacer.Replace(cellValue)
}

func main() {
	var rec fmxmlSnippet

	debug := flag.Bool("debug", false, "write debug.log and output.xml")
	flag.Parse()

	exe, err := os.Executable()
	if err != nil {
		log.Fatal(err)
	}
	dir := filepath.Dir(exe)

	if *debug {
		logFile, err := os.OpenFile(filepath.Join(dir, "debug.log"), os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
		if err != nil {
			log.Fatal(err)
		}
		defer logFile.Close()
		log.SetOutput(logFile)
	} else {
		log.SetOutput(io.Discard)
	}

	r, err := os.Open(filepath.Join(dir, "config.xml"))
	if err != nil {
		log.Fatal(err)
	}
	defer r.Close()

	if err = xml.NewDecoder(r).Decode(&rec); err != nil {
		log.Fatal(err)
	}
	xlsxFile, err := excelize.OpenFile(flag.Arg(0))
	if err != nil {
		log.Fatal(err)
	}
	defer xlsxFile.Close()

	_, rowAxis, err := excelize.SplitCellName(rec.BaseTable.Field.ID)
	if err != nil {
		log.Fatal(err)
	}

	cell := func(sheetName string, rowIndex int, cellName, defaultValue string) string {
		return returnCellValue(xlsxFile, sheetName, rowIndex, cellName, defaultValue)
	}

	rootElement := &xmlquery.Node{
		Data: "fmxmlsnippet",
		Type: xmlquery.ElementNode,
		Attr: []xmlquery.Attr{
			{Name: xml.Name{Local: "type"}, Value: "FMObjectList"},
		},
	}

	for index, sheetName := range xlsxFile.GetSheetMap() {
		fmt.Println(index, sheetName)
		if sheetName == "#SAMPLE" {
			continue
		}
		rows, err := xlsxFile.GetRows(sheetName)
		if err != nil {
			log.Println(err)
			continue
		}
		if len(rows) == 0 {
			continue
		}

		cellValue, _ := xlsxFile.GetCellValue(sheetName, rec.BaseTable.Name)
		tableElement := &xmlquery.Node{
			Data: "BaseTable",
			Type: xmlquery.ElementNode,
			Attr: []xmlquery.Attr{
				{Name: xml.Name{Local: "name"}, Value: cellValue},
			},
		}

		fieldXML := rec.BaseTable.Field
		// trailing empty cells are stripped per row, so row lengths may differ — read cell values directly from sheet
		for rowIndex, row := range rows {
			if rowIndex < rowAxis-1 || len(row) <= 1 {
				continue
			}

			fieldType := cell(sheetName, rowIndex, fieldXML.FieldType, "Normal")
			fieldElement := &xmlquery.Node{
				Data: "Field",
				Type: xmlquery.ElementNode,
				Attr: []xmlquery.Attr{
					{Name: xml.Name{Local: "id"}, Value: cell(sheetName, rowIndex, fieldXML.ID, strconv.Itoa(rowIndex))},
					{Name: xml.Name{Local: "name"}, Value: cell(sheetName, rowIndex, fieldXML.Name, fmt.Sprintf("Field#%d", rowIndex))},
					{Name: xml.Name{Local: "fieldType"}, Value: fieldType},
					{Name: xml.Name{Local: "dataType"}, Value: cell(sheetName, rowIndex, fieldXML.DataType, "Text")},
				},
			}

			commentElement := &xmlquery.Node{Data: "Comment", Type: xmlquery.ElementNode}
			xmlquery.AddChild(commentElement, &xmlquery.Node{
				Data: cell(sheetName, rowIndex, fieldXML.Comment, ""),
				Type: xmlquery.TextNode,
			})
			xmlquery.AddChild(fieldElement, commentElement)

			if fieldType == "Calculated" {
				calcElement := &xmlquery.Node{
					Data: "Calculation",
					Type: xmlquery.ElementNode,
					Attr: []xmlquery.Attr{
						{Name: xml.Name{Local: "table"}, Value: cell(sheetName, rowIndex, fieldXML.Calculation.Table, "")},
					},
				}
				xmlquery.AddChild(calcElement, &xmlquery.Node{
					Data: cell(sheetName, rowIndex, fieldXML.Calculation.Value, ""),
					Type: xmlquery.TextNode,
				})
				xmlquery.AddChild(fieldElement, calcElement)
			}

			autoEnterElement := &xmlquery.Node{
				Data: "AutoEnter",
				Type: xmlquery.ElementNode,
				Attr: []xmlquery.Attr{
					{Name: xml.Name{Local: "constant"}, Value: "False"},
					{Name: xml.Name{Local: "calculation"}, Value: "False"},
					{Name: xml.Name{Local: "alwaysEvaluate"}, Value: cell(sheetName, rowIndex, fieldXML.AutoEnter.AlwaysEvaluate, "False")},
					{Name: xml.Name{Local: "overwriteExistingValue"}, Value: cell(sheetName, rowIndex, fieldXML.AutoEnter.OverwriteExistingValue, "True")},
					{Name: xml.Name{Local: "allowEditing"}, Value: cell(sheetName, rowIndex, fieldXML.AutoEnter.AllowEditing, "False")},
					{Name: xml.Name{Local: "furigana"}, Value: cell(sheetName, rowIndex, fieldXML.AutoEnter.Furigana, "False")},
					{Name: xml.Name{Local: "lookup"}, Value: cell(sheetName, rowIndex, fieldXML.AutoEnter.Lookup, "False")},
				},
			}

			constantDataElement := &xmlquery.Node{Data: "ConstantData", Type: xmlquery.ElementNode}
			switch autoEnterConstant := cell(sheetName, rowIndex, fieldXML.AutoEnter.Constant, ""); autoEnterConstant {
			case "固定値":
				autoEnterElement.SetAttr("constant", "True")
			case "作成TS":
				autoEnterElement.SetAttr("value", "CreationTimeStamp")
			case "作成者":
				autoEnterElement.SetAttr("value", "CreationAccountName")
			case "修正TS":
				autoEnterElement.SetAttr("value", "ModificationTimeStamp")
			case "修正者":
				autoEnterElement.SetAttr("value", "ModificationAccountName")
			case "計算値":
				autoEnterElement.SetAttr("calculation", "True")
				constantDataElement = &xmlquery.Node{
					Data: "Calculation",
					Type: xmlquery.ElementNode,
					Attr: []xmlquery.Attr{
						{Name: xml.Name{Local: "table"}, Value: ""},
					},
				}
			}
			xmlquery.AddChild(constantDataElement, &xmlquery.Node{
				Data: cell(sheetName, rowIndex, fieldXML.AutoEnter.ConstantData, ""),
				Type: xmlquery.TextNode,
			})
			xmlquery.AddChild(autoEnterElement, constantDataElement)
			xmlquery.AddChild(fieldElement, autoEnterElement)

			validationElement := &xmlquery.Node{
				Data: "Validation",
				Type: xmlquery.ElementNode,
				Attr: []xmlquery.Attr{
					{Name: xml.Name{Local: "maxLength"}, Value: "False"},
					{Name: xml.Name{Local: "message"}, Value: cell(sheetName, rowIndex, fieldXML.Validation.Message, "False")},
					{Name: xml.Name{Local: "valuelist"}, Value: cell(sheetName, rowIndex, fieldXML.Validation.Valuelist, "False")},
					{Name: xml.Name{Local: "calculation"}, Value: cell(sheetName, rowIndex, fieldXML.Validation.Calculation, "False")},
					{Name: xml.Name{Local: "alwaysValidateCalculation"}, Value: cell(sheetName, rowIndex, fieldXML.Validation.AlwaysValidateCalculation, "False")},
					{Name: xml.Name{Local: "type"}, Value: cell(sheetName, rowIndex, "", "OnlyDuringDataEntry")},
				},
			}

			xmlquery.AddChild(validationElement, &xmlquery.Node{
				Data: "Unique",
				Type: xmlquery.ElementNode,
				Attr: []xmlquery.Attr{
					{Name: xml.Name{Local: "value"}, Value: cell(sheetName, rowIndex, fieldXML.Validation.Unique.Value, "True")},
				},
			})

			strictValidationElement := &xmlquery.Node{
				Data: "StrictValidation",
				Type: xmlquery.ElementNode,
				Attr: []xmlquery.Attr{{Name: xml.Name{Local: "value"}, Value: "False"}},
			}
			if strictDataTypeValue := cell(sheetName, rowIndex, fieldXML.Validation.StrictDataType.Value, ""); strictDataTypeValue != "" {
				strictValidationElement.SetAttr("value", "True")
				xmlquery.AddChild(validationElement, &xmlquery.Node{
					Data: "StrictDataType",
					Type: xmlquery.ElementNode,
					Attr: []xmlquery.Attr{{Name: xml.Name{Local: "value"}, Value: strictDataTypeValue}},
				})
			}
			xmlquery.AddChild(validationElement, strictValidationElement)

			xmlquery.AddChild(validationElement, &xmlquery.Node{
				Data: "NotEmpty",
				Type: xmlquery.ElementNode,
				Attr: []xmlquery.Attr{
					{Name: xml.Name{Local: "value"}, Value: cell(sheetName, rowIndex, fieldXML.Validation.NotEmpty.Value, "True")},
				},
			})

			maxLengthValue := cell(sheetName, rowIndex, fieldXML.Validation.MaxDataLength.Value, "")
			if maxLengthValue != "" {
				validationElement.SetAttr("maxLength", "True")
			}
			xmlquery.AddChild(validationElement, &xmlquery.Node{
				Data: "MaxDataLength",
				Type: xmlquery.ElementNode,
				Attr: []xmlquery.Attr{{Name: xml.Name{Local: "value"}, Value: maxLengthValue}},
			})
			xmlquery.AddChild(fieldElement, validationElement)

			xmlquery.AddChild(fieldElement, &xmlquery.Node{
				Data: "Storage",
				Type: xmlquery.ElementNode,
				Attr: []xmlquery.Attr{
					{Name: xml.Name{Local: "autoIndex"}, Value: cell(sheetName, rowIndex, fieldXML.Storage.AutoIndex, "True")},
					{Name: xml.Name{Local: "index"}, Value: cell(sheetName, rowIndex, fieldXML.Storage.Index, "None")},
					{Name: xml.Name{Local: "indexLanguage"}, Value: cell(sheetName, rowIndex, fieldXML.Storage.IndexLanguage, "Japanese")},
					{Name: xml.Name{Local: "global"}, Value: cell(sheetName, rowIndex, fieldXML.Storage.Global, "False")},
					{Name: xml.Name{Local: "maxRepetition"}, Value: cell(sheetName, rowIndex, fieldXML.Storage.MaxRepetition, "1")},
				},
			})
			xmlquery.AddChild(tableElement, fieldElement)
		}
		xmlquery.AddChild(rootElement, tableElement)
	}

	xmlStr := rootElement.OutputXML(true)

	if *debug {
		if err = os.WriteFile(filepath.Join(dir, "output.xml"), []byte(prettyXML(xmlStr)), 0644); err != nil {
			log.Println(err)
		}
	}

	switch runtime.GOOS {
	case "darwin":
		// https://stackoverflow.com/questions/45248144
		if err = exec.Command("/usr/bin/osascript", "-e",
			fmt.Sprintf(`set the clipboard to «data XMTB%s»`, hex.EncodeToString([]byte(xmlStr))),
		).Run(); err != nil {
			log.Println(err)
		}
	case "windows":
		if err = clipboard.WriteAll(xmlStr); err != nil {
			log.Println(err)
		}
	default:
		log.Println("unsupported OS")
	}
}
