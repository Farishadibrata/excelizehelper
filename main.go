package main

import (
	"fmt"
	"time"

	"github.com/xuri/excelize/v2"
)

type ExcelizeInstance struct {
	Excelize      *excelize.File
	SheetName     string
	CurrentCoords *Coords
	TableSpacing  int
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

// To appending Table from previous to avoid overlapping
func (eInstance *ExcelizeInstance) ReFetchCoords() {
	rows, err := eInstance.Excelize.Rows(eInstance.SheetName)
	if err != nil {
		fmt.Println(err)
		return
	}
	spacingBetweenTable := eInstance.TableSpacing
	maxColLength := 0
	for rows.Next() {
		cols, err := rows.Columns()
		if err != nil {
			fmt.Println(err)
		}
		if len(cols) > maxColLength {
			maxColLength = len(cols)
		}
	}
	if maxColLength == 0 {
		spacingBetweenTable = 0
	}
	eInstance.CurrentCoords.setCoordsX(eInstance.CurrentCoords.X + maxColLength + spacingBetweenTable)

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
	Style int
	// Merging X / Z based on Current X / Z + MergeX / MergeZ
	MergeY int
	MergeX int
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
}

type ITable struct {
	Rows []IRows
	// Automatically add filter after first row (assuming first row is header)
	AutoFilter bool
}

type INewExcelInstance struct {
	sheetName string
}

func (eInstance *ExcelizeInstance) Write() {
	fileName := time.Now().Format("20060102150405")
	if err := eInstance.Excelize.SaveAs(fmt.Sprintf("%s.xlsx", fileName)); err != nil {
		fmt.Println(err)
	}
}

func (eInstance *ExcelizeInstance) AppendTable(input *ITable) {
	//  add spacing between table
	if eInstance.CurrentCoords.X < 1 {
		eInstance.CurrentCoords.X = eInstance.CurrentCoords.X + eInstance.TableSpacing + 1
	}

	eInstance.ReFetchCoords()
	tableCoords := &Coords{
		X: eInstance.CurrentCoords.X,
		Y: eInstance.CurrentCoords.Y,
	}

	Headerindex := 1
	// if there are previously added newline in autofilter, skip it
	for index, row := range input.Rows {
		isColumnRenderable := true

		if row.Header {
			Headerindex = index + 1
		}
		if isColumnRenderable {
			for _, column := range row.Columns {
				currentCoords := tableCoords.currentCoordsToCell()
				eInstance.Excelize.SetCellValue(eInstance.SheetName, currentCoords, column.V)

				if column.MergeY != 0 {
					mergedCell := tableCoords.currentCellPlusN(0, column.MergeY-1)
					eInstance.Excelize.MergeCell(eInstance.SheetName, currentCoords, mergedCell)
				}

				if column.MergeX != 0 {
					mergedCell := tableCoords.currentCellPlusN(column.MergeX-1, 0)
					eInstance.Excelize.MergeCell(eInstance.SheetName, currentCoords, mergedCell)
				}

				if row.OutlineX.N != 0 && row.OutlineX.Level != 0 {
					eInstance.Excelize.SetColOutlineLevel(eInstance.SheetName, getColumnName(tableCoords.X+column.OutlineX.N), 1)
				}

				tableCoords.currentCoordsAddCol(column.MergeX)
			}
		}

		if row.OutlineY.N != 0 && row.OutlineY.Level != 0 {
			eInstance.Excelize.SetRowOutlineLevel(eInstance.SheetName, tableCoords.Y+row.OutlineY.N, 1)
		}

		tableCoords.setCoordsX(eInstance.CurrentCoords.X)
		tableCoords.currentCoordsAddRow()
	}
	// Always on header index
	if input.AutoFilter {
		HeaderLength := len(input.Rows[Headerindex].Columns)
		rowLength := len(input.Rows)
		autoFilterIndexRow := tableCoords.Y - rowLength + Headerindex - 1

		startRange, _ := excelize.CoordinatesToCellName(eInstance.CurrentCoords.X, autoFilterIndexRow)
		endRange, _ := excelize.CoordinatesToCellName(eInstance.CurrentCoords.X+HeaderLength-1, autoFilterIndexRow)
		rangeAutoFilter := fmt.Sprintf("%s:%s", startRange, endRange)

		eInstance.Excelize.AutoFilter(eInstance.SheetName, rangeAutoFilter, []excelize.AutoFilterOptions{})
	}

}

func NewExcelInstance(input *INewExcelInstance) *ExcelizeInstance {
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// Create a new sheet.
	_, err := f.NewSheet(input.sheetName)

	if err != nil {
		fmt.Println(err)
		panic("Boom")
		// return nil
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
	return instance
}

func main() {

	xlsx := NewExcelInstance(&INewExcelInstance{sheetName: "Sheet1"})

	Headers := []IColumns{{V: "Supplier Name"}, {V: "PO Number"}, {V: "Discipline"}}

	MainTable := &ITable{
		// AutoFilter: true,
		Rows: []IRows{
			{
				Columns: []IColumns{{V: "Testing Table", BaseProps: BaseProps{MergeX: 3}}},
			},
			{
				Columns: Headers,
				Header:  true,
			},
			{
				OutlineY: Outline{N: 3},
				Columns:  []IColumns{{V: "STEEL WORLD CO., LTD"}, {V: "3220000481"}, {V: ""}},
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

	SecondTable := &ITable{
		AutoFilter: true,
		Rows: []IRows{
			{
				Columns: Headers,
				Header:  true,
			},
			{
				OutlineY: Outline{N: 3},
				Columns:  []IColumns{{V: "STEEL WORLD CO., LTD"}, {V: "3220000481"}, {V: ""}},
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

	xlsx.AppendTable(MainTable)
	xlsx.AppendTable(SecondTable)

	xlsx.Write()
}
