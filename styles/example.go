package styles

import "github.com/xuri/excelize/v2"

var CenteringBoldStyle = excelize.Style{
	Font: &excelize.Font{Bold: true},
	Alignment: &excelize.Alignment{
		Horizontal: "center",
	},
}

var CenteringWidthStyle = excelize.Style{
	Alignment: &excelize.Alignment{
		Horizontal: "center",
	},
}

var TableHeaderColumnsStyle = excelize.Style{
	Alignment: &excelize.Alignment{
		Horizontal:  "center",
		WrapText:    true,
		ShrinkToFit: true,
		Vertical:    "center",
	},
	Fill: excelize.Fill{
		Type:    "pattern",
		Color:   []string{"#c6efce"},
		Pattern: 1,
	},
	Border: []excelize.Border{{Type: "left", Color: "#000000", Style: 1}, {Type: "right", Color: "#000000", Style: 1}, {Type: "top", Color: "#000000", Style: 1}, {Type: "bottom", Color: "#000000", Style: 1}},
}

var TableHeaderIssuePurposeColumnsStyle = excelize.Style{
	Alignment: &excelize.Alignment{
		Horizontal:  "center",
		WrapText:    true,
		ShrinkToFit: true,
		Vertical:    "center",
	},
	Fill: excelize.Fill{
		Type:    "pattern",
		Color:   []string{"#f5e798"},
		Pattern: 1,
	},
	Border: []excelize.Border{{Type: "left", Color: "#000000", Style: 1}, {Type: "right", Color: "#000000", Style: 1}, {Type: "top", Color: "#000000", Style: 1}, {Type: "bottom", Color: "#000000", Style: 1}},
}

var TableHeaderLatestStatusColumnsStyle = excelize.Style{
	Alignment: &excelize.Alignment{
		Horizontal:  "center",
		WrapText:    true,
		ShrinkToFit: true,
		Vertical:    "center",
	},
	Fill: excelize.Fill{
		Type:    "pattern",
		Color:   []string{"#efc6d7"},
		Pattern: 1,
	},
	Border: []excelize.Border{{Type: "left", Color: "#000000", Style: 1}, {Type: "right", Color: "#000000", Style: 1}, {Type: "top", Color: "#000000", Style: 1}, {Type: "bottom", Color: "#000000", Style: 1}},
}

var EachMetadataStyle = excelize.Style{
	Fill: excelize.Fill{
		Type:    "pattern",
		Color:   []string{"#BDD7EE"},
		Pattern: 1,
	},
	Border: []excelize.Border{{Type: "left", Color: "#000000", Style: 1}, {Type: "right", Color: "#000000", Style: 1}, {Type: "top", Color: "#000000", Style: 1}, {Type: "bottom", Color: "#000000", Style: 1}},
	Font:   &excelize.Font{Bold: true},
}
