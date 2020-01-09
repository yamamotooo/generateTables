package main

import (
	"encoding/hex"
	"fmt"
	"os/exec"
	"strings"

	"os"
	"path/filepath"

	"github.com/tealeg/xlsx" //go get github.com/tealeg/xlsx
	"runtime"

	"./clipboard"
)

var (
	snippetXML	= `<fmxmlsnippet type="FMObjectList">%s</fmxmlsnippet>`
	tableXML	= `<BaseTable name="%s">%s</BaseTable>`
	fieldXML	= `<Field id="%s" dataType="%s" fieldType="%s" name="%s">
					<Comment>%s
					</Comment>
					<AutoEnter allowEditing="%s" constant="%s" furigana="%s" lookup="%s" calculation="%s">
					<ConstantData>%s
					</ConstantData>
					</AutoEnter>
					<Validation message="%s" maxLength="%s" valuelist="%s" calculation="%s" alwaysValidateCalculation="%s" type="%s">
					<NotEmpty value="%s">
					</NotEmpty>
					<Unique value="%s">
					</Unique>
					<Existing value="%s">
					</Existing>
					<StrictValidation value="%s">
					</StrictValidation>
					</Validation>
					<Storage autoIndex="%s" index="%s" indexLanguage="%s" global="%s" maxRepetition="%s">
					</Storage>
				</Field>`
)

func main() {

	//xlsxPath := "your file path here or use below"
	xlsxPath, err := os.Executable()
	//xlsxPath, err := os.Getwd()
	xlsxPath = filepath.Dir(xlsxPath) + "/Book1.xlsx"
	//fmt.Println(xlsxPath)
	
	xlsxFile, err := xlsx.OpenFile(xlsxPath)
	if err != nil {
		panic("Excel file not found...")
	}

	tables := ""
	for _, sheet := range xlsxFile.Sheets {
		fields := ""
		for i := 3; i < sheet.MaxRow; i++ {
			fieldTemp := fieldXML
			row := sheet.Row(i)
			for _, cell := range row.Cells {
				fieldTemp = strings.Replace(fieldTemp, "%s", cell.String(), 1)
			}
			fields += fieldTemp
		}
		tables += fmt.Sprintf(tableXML, sheet.Rows[0].Cells[1].String(), fields)
	}

	snippetXML = fmt.Sprintf(snippetXML, tables)
	snippetXML = strings.NewReplacer("\t", "", "\r\n", "", "\r", "", "\n", "").Replace(string(snippetXML))

	switch runtime.GOOS {
	case "darwin":
		hexData := hex.EncodeToString([]byte(snippetXML))
		hexData = "XMTB" + hexData
		hexData = "«data " + hexData + "»"
		err = exec.Command("osascript", "-e", fmt.Sprintf("set the clipboard to %s", hexData)).Run()
		if err != nil {
			panic(err)
		}
		//fallthrough 次節へ処理を流す場合
	case "windows":
		clipboard.WriteAll(snippetXML)
		//text, _ := clipboard.ReadAll()
		//fmt.Println(text)
		//fallthrough 次節へ処理を流す場合
	default:
		fmt.Println("ohter")
	}

}

//GOOS=darwin GOARCH=amd64 go build generateTables.go
//GOOS=windows GOARCH=amd64 go build generateTables.go




