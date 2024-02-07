package main

import (
	"fmt"
	"time"

	"farishadibrata.com/xlsxparser/styles"
	"github.com/xuri/excelize/v2"
)

type ExcelizeInstance struct {
	Excelize      *excelize.File
	SheetName     string
	CurrentCoords *Coords
	TableSpacing  int
	Style         map[string]int
	Debug         bool
}

type IExcelizeStyle struct {
	Name          string
	ExcelizeStyle excelize.Style
}

// https://stackoverflow.com/a/71354018
func getColumnName(col int) string {
	name := make([]byte, 0, 3) // max 16,384 columns (2022)
	const aLen = 'Z' - 'A' + 1 // alphabet length
	for ; col > 0; col /= aLen + 1 {
		name = append(name, byte('A'+(col-1)%aLen))
	}
	for i, j := 0, len(name)-1; i < j; i, j = i+1, j-1 {
		name[i], name[j] = name[j], name[i]
	}
	return string(name)
}

// This function to cycle between color that can be use as different color for each table
func ColorCycle(index int) string {
	ColorList := []string{"#ffb3ba", "#ffdfba", "#ffffba", "#baffc9", "#bae1ff"}
	return ColorList[index%len(ColorList)]
}

type Coords struct {
	X int
	Y int
}

func (coords *Coords) currentCoordsToCell() string {
	cell, _ := excelize.CoordinatesToCellName(coords.X, coords.Y)
	return cell
}

func (coords *Coords) currentCellPlusN(x, y int) string {
	cell, _ := excelize.CoordinatesToCellName(coords.X+x, coords.Y+y)
	return cell
}
func (coords *Coords) currentCoordsAddRow() {
	coords.Y = coords.Y + 1
}
func (coords *Coords) setCoordsX(n int) {
	coords.X = n
}
func (coords *Coords) setCoordsY(n int) {
	coords.Y = n
}

func (coords *Coords) currentCoordsAddCol(n int) {
	if n == 0 {
		n = 1
	}
	coords.X = coords.X + n
}

type Outline struct {
	N     int
	Level int
}
type BaseProps struct {
	// Registered style id
	Style string
	// Merging X / Z based on Current X / Z + MergeX / MergeZ
	MergeY int
	MergeX int
	Width  int
	// Outlining current row + n
	OutlineX Outline
}
type IColumns struct {
	// V means values for shorten code
	BaseProps
	V string
}

type IRows struct {
	BaseProps
	Columns  []IColumns
	OutlineY Outline
	Header   bool
	Height   int
	EmptyRow bool
}

type ITable struct {
	Rows []IRows
	// Automatically add filter after first row (assuming first row is header)
	AutoFilter bool
}

type INewExcelInstance struct {
	sheetName string
}

func (eInstance *ExcelizeInstance) Log(message string, errorMsg error) {
	if eInstance.Debug {
		println(fmt.Errorf("Excelize-Helper: %s: %v", message, errorMsg).Error())
	}
}

// To appending Table from previous to avoid overlapping
func (eInstance *ExcelizeInstance) ReFetchCoords() error {
	rows, err := eInstance.Excelize.Rows(eInstance.SheetName)
	if err != nil {
		return err
	}
	spacingBetweenTable := eInstance.TableSpacing
	maxColLength := 0
	for rows.Next() {
		cols, err := rows.Columns()
		if err != nil {
			return err
		}
		if len(cols) > maxColLength {
			maxColLength = len(cols)
		}
	}
	if maxColLength == 0 {
		spacingBetweenTable = 0
	}
	eInstance.CurrentCoords.setCoordsX(eInstance.CurrentCoords.X + maxColLength + spacingBetweenTable)
	return nil
}
func (eInstance *ExcelizeInstance) AppendStyle(Styles []IExcelizeStyle) error {
	newStyles := make(map[string]int, len(Styles))
	for _, style := range Styles {
		newStyle, err := eInstance.Excelize.NewStyle(&style.ExcelizeStyle)
		if err != nil {
			eInstance.Log("Unable to save style: ", err)
			return err
		}
		newStyles[style.Name] = newStyle
	}
	eInstance.Style = newStyles
	return nil
}

func (eInstance *ExcelizeInstance) Write() error {
	fileName := time.Now().Format("20060102150405")
	if err := eInstance.Excelize.SaveAs(fmt.Sprintf("%s.xlsx", fileName)); err != nil {
		eInstance.Log("Unable to save XLSX: ", err)
		return err
	}
	return nil
}

func (eInstance *ExcelizeInstance) AppendTable(input *ITable) error {
	//  add spacing between table
	if eInstance.CurrentCoords.X > 1 {
		eInstance.CurrentCoords.X = eInstance.CurrentCoords.X + eInstance.TableSpacing + 1
	}

	errRefetch := eInstance.ReFetchCoords()
	if errRefetch != nil {
		eInstance.Log("Unable to save style: ", errRefetch)
		return errRefetch
	}

	tableCoords := &Coords{
		X: eInstance.CurrentCoords.X,
		Y: eInstance.CurrentCoords.Y,
	}

	Headerindex := 1
	// if there are previously added newline in autofilter, skip it
	for index, row := range input.Rows {
		isColumnRenderable := true

		if row.EmptyRow {
			isColumnRenderable = false
		}
		if row.Header {
			Headerindex = index + 1
		}
		if isColumnRenderable {
			for _, column := range row.Columns {
				currentCoords := tableCoords.currentCoordsToCell()
				eInstance.Excelize.SetCellValue(eInstance.SheetName, currentCoords, column.V)

				if column.Width != 0 {
					eInstance.Excelize.SetColWidth(eInstance.SheetName, getColumnName(tableCoords.X), getColumnName(tableCoords.X), float64(column.Width))
				}

				if column.Style != "" {
					eInstance.Excelize.SetCellStyle(eInstance.SheetName, currentCoords, currentCoords, eInstance.Style[column.Style])
				}

				if column.MergeY != 0 {
					mergedCell := tableCoords.currentCellPlusN(0, column.MergeY-1)
					eInstance.Excelize.MergeCell(eInstance.SheetName, currentCoords, mergedCell)
				}

				if column.MergeX != 0 {
					mergedCell := tableCoords.currentCellPlusN(column.MergeX-1, 0)
					eInstance.Excelize.MergeCell(eInstance.SheetName, currentCoords, mergedCell)
				}

				if row.OutlineX.N != 0 && row.OutlineX.Level != 0 {
					eInstance.Excelize.SetColOutlineLevel(eInstance.SheetName, getColumnName(tableCoords.X), 1)
				}

				tableCoords.currentCoordsAddCol(column.MergeX)
			}
		}

		if row.OutlineY.Level != 0 {
			eInstance.Excelize.SetRowOutlineLevel(eInstance.SheetName, tableCoords.Y, 1)
		}

		if row.Height != 0 {
			eInstance.Excelize.SetRowHeight(eInstance.SheetName, index+1, float64(row.Height))
		}

		if row.Style != "" {
			//BUG: it will overlap with each other if there is already style applied before
			eInstance.Excelize.SetRowStyle(eInstance.SheetName, index+1, index+1, eInstance.Style[row.Style])
		}

		tableCoords.setCoordsX(eInstance.CurrentCoords.X)
		tableCoords.currentCoordsAddRow()
	}
	//NOTE: AutoFilter Always on header index and only one per sheet for now
	if input.AutoFilter {
		HeaderLength := len(input.Rows[Headerindex].Columns)
		rowLength := len(input.Rows)
		autoFilterIndexRow := tableCoords.Y - rowLength + Headerindex - 1

		startRange, _ := excelize.CoordinatesToCellName(eInstance.CurrentCoords.X, autoFilterIndexRow)
		endRange, _ := excelize.CoordinatesToCellName(eInstance.CurrentCoords.X+HeaderLength-1, autoFilterIndexRow)
		rangeAutoFilter := fmt.Sprintf("%s:%s", startRange, endRange)

		eInstance.Excelize.AutoFilter(eInstance.SheetName, rangeAutoFilter, []excelize.AutoFilterOptions{})
	}
	return nil
}

func StringToArrayColumns(baseProps BaseProps, input ...string) []IColumns {
	columns := make([]IColumns, len(input))
	for i, v := range input {
		columns[i] = IColumns{
			BaseProps: baseProps,
			V:         string(v),
		}
	}
	return columns
}

func NewExcelInstance(input *INewExcelInstance) (*ExcelizeInstance, error) {
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// Create a new sheet.
	_, err := f.NewSheet(input.sheetName)

	if err != nil {
		fmt.Println("Excelize-Helper: Unable to create sheet: ", err)
		return nil, err
	}

	instance := &ExcelizeInstance{
		Excelize:  f,
		SheetName: input.sheetName,
		// excel coords start at 1
		CurrentCoords: &Coords{
			X: 1,
			Y: 1,
		},
		TableSpacing: 1,
	}
	return instance, nil
}

func main() {

	xlsx, _ := NewExcelInstance(&INewExcelInstance{sheetName: "Sheet1"})

	xlsx.AppendStyle([]IExcelizeStyle{
		{Name: "CenteringBoldStyle", ExcelizeStyle: styles.CenteringBoldStyle},
		{Name: "CenteringWidthStyle", ExcelizeStyle: styles.CenteringWidthStyle},
		{Name: "TableHeaderColumnsStyle", ExcelizeStyle: styles.TableHeaderColumnsStyle},
		{Name: "TableHeaderIssuePurposeColumnsStyle", ExcelizeStyle: styles.TableHeaderIssuePurposeColumnsStyle},
		{Name: "TableHeaderLatestStatusColumnsStyle", ExcelizeStyle: styles.TableHeaderLatestStatusColumnsStyle},
		{Name: "EachMetadataStyle", ExcelizeStyle: styles.EachMetadataStyle},
	})

	HeaderStyle := BaseProps{
		Style: "TableHeaderColumnsStyle",
		Width: 15,
	}
	Headers := StringToArrayColumns(HeaderStyle, "Supplier Name", "PO Number", "PO TITLE", "CTR DOC. NO.", "CPY DOC. NO.", "DISTRIBUTION MATRIX", "DOCUMENT TITLE", "DOC TYPE", "PIC DCC VENDOR", "PIC EXPEDITER", "PIC QAQC", "Vendor Plan Date for 1st Submission", "CTR", "CPY", "FINAL ISSUANCE", "REMARKS")
	HeadersIssuePurpose := StringToArrayColumns(HeaderStyle, "Plan Date", "Vendor Minor Revision", "Vendor Major Revision	Status", "Actual Date", "Vendor Late / On Schedule", "CTR Review Date", "Due Date", "CTR Return Date", "Approval Code", "Return DOC To Vendor", "CTR Late / On Schedule", "TRN TRIPATRA OUT", "Expected Returned To TPEC")

	// HeadersLatestStatus := []IColumns{{V: "Plan Date"}, {V: "Vendor Minor Revision"}, {V: "Vendor Major Revision"}}

	MetadataTable := &ITable{
		AutoFilter: true,
		Rows: []IRows{
			{EmptyRow: true},
			{
				Columns: Headers,
			},
			{
				Columns: []IColumns{},
				Header:  true,
			},
			{
				BaseProps: BaseProps{Style: "EachMetadataStyle"},
				Columns:   []IColumns{{V: "STEEL WORLD CO., LTD"}, {V: "3220000481"}, {V: ""}},
			},
			{
				OutlineY: Outline{Level: 1},
				Columns:  []IColumns{{V: "STEEL WORLD CO., LTD"}, {V: "3220000481"}, {V: "Piping"}},
			},
			{
				OutlineY: Outline{Level: 1},
				Columns:  []IColumns{{V: "STEEL WORLD CO., LTD"}, {V: "3220000481"}, {V: "Electrical"}},
			},
			{
				OutlineY: Outline{Level: 1},
				Columns:  []IColumns{{V: "STEEL WORLD CO., LTD"}, {V: "3220000481"}, {V: "Mechanical"}},
			},
		},
	}

	IssuePurposeTable := &ITable{
		// AutoFilter: true,
		Rows: []IRows{
			{
				Columns: []IColumns{{V: "Latest Status", BaseProps: BaseProps{Style: "CenteringBoldStyle", MergeX: len(HeadersIssuePurpose)}}},
			},
			{
				Height:  50,
				Columns: HeadersIssuePurpose,
				Header:  true,
			},
			{
				Columns: []IColumns{{V: "STEEL WORLD CO., LTD"}, {V: "3220000481"}, {V: ""}},
			},
			{
				Columns: []IColumns{{V: "STEEL WORLD CO., LTD"}, {V: "3220000481"}, {V: "Piping"}},
			},
			{
				Columns: []IColumns{{V: "STEEL WORLD CO., LTD"}, {V: "3220000481"}, {V: "Electrical"}},
			},
			{
				Columns: []IColumns{{V: "STEEL WORLD CO., LTD"}, {V: "3220000481"}, {V: "Mechanical"}},
			},
		},
	}

	xlsx.AppendTable(MetadataTable)
	xlsx.AppendTable(IssuePurposeTable)
	xlsx.AppendTable(IssuePurposeTable)
	xlsx.Excelize.SetPanes(xlsx.SheetName, &excelize.Panes{
		Freeze:      true,
		XSplit:      8,
		YSplit:      1,
		TopLeftCell: "I2",
	})

	xlsx.Write()
}
