package main

import (
	"encoding/hex"
	"fmt"
	"os/exec"
	"strings"

	"github.com/tealeg/xlsx"
)

var snippetXML = `<fmxmlsnippet type="FMObjectList">%s</fmxmlsnippet>`
var tableXML = `<BaseTable name="%s">%s</BaseTable>`
var fieldXML = `<Field id="%s" dataType="%s" fieldType="%s" name="%s">
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

func main() {

	xlsxPath := "your file path here or use below"
	//xlsxPath, err := os.Executable()
	//xlsxPath = filepath.Dir(xlsxPath) + "/Book1.xlsx"

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

	hexData := hex.EncodeToString([]byte(snippetXML))
	hexData = "XMTB" + hexData
	hexData = "«data " + hexData + "»"
	/*
		Mac-XMTB = table
		Mac-XMFD = field
		Mac-XMSC = script
		Mac-XMSS = script step
		Mac-XMFN = custom function
		Mac-XMLO = layout object (.fp7)
		Mac-XML2 = layout object (.fmp12)
		Mac-XMVL = value list (FM16)
		Mac-     = Theme
	*/

	err = exec.Command("osascript", "-e", fmt.Sprintf("set the clipboard to %s", hexData)).Run()
	if err != nil {
		panic(err)
	}

}
