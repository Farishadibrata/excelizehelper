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

type BaseProps struct {
	// Registered style id
	Style int
	// Merging X / Z based on Current X / Z + MergeX / MergeZ
	MergeY int
	MergeX int
	// current X & y coords
}
type IColumns struct {
	BaseProps
	Value string
}

type IRows struct {
	BaseProps
	Columns []IColumns
}

type ITable struct {
	Rows []IRows
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

	for _, row := range input.Rows {
		for _, column := range row.Columns {

			currentCoords := tableCoords.currentCoordsToCell()
			eInstance.Excelize.SetCellValue(eInstance.SheetName, currentCoords, column.Value)

			if column.MergeY != 0 {
				mergedCell := tableCoords.currentCellPlusN(0, column.MergeY-1)
				eInstance.Excelize.MergeCell(eInstance.SheetName, currentCoords, mergedCell)
			}

			if column.MergeX != 0 {
				mergedCell := tableCoords.currentCellPlusN(column.MergeX-1, 0)
				eInstance.Excelize.MergeCell(eInstance.SheetName, currentCoords, mergedCell)
			}
			tableCoords.currentCoordsAddCol(column.MergeX)
		}
		tableCoords.setCoordsX(eInstance.CurrentCoords.X)
		tableCoords.currentCoordsAddRow()
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

	MainTable := &ITable{
		Rows: []IRows{{
			Columns: []IColumns{{Value: "Properties", BaseProps: BaseProps{
				MergeX: 3,
			}}, {Value: "Departement", BaseProps: BaseProps{
				MergeY: 2,
			}}},
		}, {
			Columns: []IColumns{{Value: "Name"}, {Value: "Task"}, {Value: "Assigned"}},
		}, {
			Columns: []IColumns{{Value: "Nora"}, {Value: "Write Docs"}, {Value: "false"}, {Value: "Business Analyst"}},
		},
			{
				Columns: []IColumns{{Value: "Tyler"}, {Value: "Write Code for FrontEnd"}, {Value: "false"}, {Value: "IT Departement"}},
			},
			{
				Columns: []IColumns{{Value: "Durden"}, {Value: "Write Code for BackEnd"}, {Value: "true"}, {Value: "IT Departement"}},
			},
		},
	}

	xlsx.AppendTable(MainTable)
	xlsx.AppendTable(MainTable)

	xlsx.Write()
}
