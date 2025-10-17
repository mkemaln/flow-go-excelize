package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func main() {
	f := excelize.NewFile()

	// the cell size is 64x20 (default by Devin AI)
	// or
	// the cell size is 64x18 (in libre)

	// this is how to center the shape with the based is the cell 64x20 is the default cell size in pixel
	// (64 - x) x is the width of the shape
	// (64 - y) y is the height of the shape

	// if we disable the centerOffsetY it will center only the X part,
	// the shape top will be render in the exact cell point
	centerOffsetX := (64 - 1) / 2
	// centerOffsetY := (20 - 80) / 2

	// baseArrowW := 40
	// baseArrowH := 80

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	lineWidth := 1.2
	f.AddShape("Sheet1",
		&excelize.Shape{
			Cell: "H2",
			Type: "ellipse",
			Line: excelize.ShapeLine{Color: "060270", Width: &lineWidth},
			Fill: excelize.Fill{Color: []string{"FFFFFF"}, Pattern: 1},
			// Paragraph: []excelize.RichTextRun{
			// 	{
			// 		Text: "Start",
			// 		Font: &excelize.Font{
			// 			Bold:      false,
			// 			Italic:    false,
			// 			Family:    "Times New Roman",
			// 			Size:      14,
			// 			Color:     "777777",
			// 			VertAlign: "superscript",
			// 		},
			// 	},
			// },
			Width:  64,
			Height: 18,
		},
	)
	f.AddShape("Sheet1",
		&excelize.Shape{
			Cell:   "H3",
			Type:   "downArrow",
			Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
			Fill:   excelize.Fill{Color: []string{"FFFFFF"}, Pattern: 1},
			Width:  1,
			Height: 40,
			Format: excelize.GraphicOptions{
				OffsetX: centerOffsetX,
			},
		},
	)
	f.AddShape("Sheet1",
		&excelize.Shape{
			Cell: "A2",
			Type: "ellipse",
			Line: excelize.ShapeLine{Color: "060270", Width: &lineWidth},
			Fill: excelize.Fill{Color: []string{"FFFFFF"}, Pattern: 1},
			// Paragraph: []excelize.RichTextRun{
			// 	{
			// 		Text: "Start",
			// 		Font: &excelize.Font{
			// 			Bold:      false,
			// 			Italic:    false,
			// 			Family:    "Times New Roman",
			// 			Size:      14,
			// 			Color:     "777777",
			// 			VertAlign: "superscript",
			// 		},
			// 	},
			// },
			Width:  64,
			Height: 18,
		},
	)
	f.AddShape("Sheet1",
		&excelize.Shape{
			Cell:   "A3",
			Type:   "line",
			Line:   excelize.ShapeLine{Color: "060270"},
			Width:  1,
			Height: 40,
			Format: excelize.GraphicOptions{
				OffsetX: centerOffsetX,
			},
		},
	)

	f.AddShape("Sheet1",
		&excelize.Shape{
			Cell:   "I6",
			Type:   "downArrow",
			Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
			Fill:   excelize.Fill{Color: []string{"FFFFFF"}, Pattern: 1},
			Width:  1,
			Height: 40,
		},
	)
	f.AddShape("Sheet1",
		&excelize.Shape{
			Cell:   "J6",
			Type:   "downArrow",
			Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
			Fill:   excelize.Fill{Color: []string{"FFFFFF"}, Pattern: 1},
			Width:  2,
			Height: 40,
		},
	)
	f.AddShape("Sheet1",
		&excelize.Shape{
			Cell:   "K6",
			Type:   "downArrow",
			Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
			Fill:   excelize.Fill{Color: []string{"FFFFFF"}, Pattern: 1},
			Width:  3,
			Height: 40,
		},
	)
	f.AddShape("Sheet1",
		&excelize.Shape{
			Cell:   "L6",
			Type:   "downArrow",
			Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
			Fill:   excelize.Fill{Color: []string{"FFFFFF"}, Pattern: 1},
			Width:  4,
			Height: 40,
		},
	)
	f.AddShape("Sheet1",
		&excelize.Shape{
			Cell: "H11",
			Type: "parallelogram",
			Line: excelize.ShapeLine{Color: "060270", Width: &lineWidth},
			Fill: excelize.Fill{Color: []string{"FFFFFF"}, Pattern: 1},
			Paragraph: []excelize.RichTextRun{
				{
					Text: "Input 1",
					Font: &excelize.Font{
						Bold:      false,
						Italic:    false,
						Family:    "Times New Roman",
						Size:      14,
						Color:     "777777",
						VertAlign: "subscript",
					},
				},
			},
			Width:  180,
			Height: 40,
		},
	)
	f.AddShape("Sheet1",
		&excelize.Shape{
			Cell:   "I14",
			Type:   "downArrow",
			Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
			Fill:   excelize.Fill{Color: []string{"FFFFFF"}, Pattern: 1},
			Width:  40,
			Height: 80,
		},
	)
	f.AddShape("Sheet1",
		&excelize.Shape{
			Cell: "I19",
			Type: "diamond",
			Line: excelize.ShapeLine{Color: "060270", Width: &lineWidth},
			Fill: excelize.Fill{Color: []string{"FFFFFF"}, Pattern: 1},
			Paragraph: []excelize.RichTextRun{
				{
					Text: "If 1",
					Font: &excelize.Font{
						Bold:      false,
						Italic:    false,
						Family:    "Times New Roman",
						Size:      14,
						Color:     "777777",
						VertAlign: "subscript",
					},
				},
			},
			Width:  80,
			Height: 80,
			Format: excelize.GraphicOptions{
				OffsetX: centerOffsetX,
				// OffsetY: centerOffsetY,
			},
		},
	)
	f.AddShape("Sheet1",
		&excelize.Shape{
			Cell:   "I23",
			Type:   "downArrow",
			Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
			Fill:   excelize.Fill{Color: []string{"FFFFFF"}, Pattern: 1},
			Width:  40,
			Height: 80,
		},
	)
	f.AddShape("Sheet1",
		&excelize.Shape{
			Cell: "H27",
			Type: "rect",
			Line: excelize.ShapeLine{Color: "060270", Width: &lineWidth},
			Fill: excelize.Fill{Color: []string{"FFFFFF"}, Pattern: 1},
			Paragraph: []excelize.RichTextRun{
				{
					Text: "Process False",
					Font: &excelize.Font{
						Bold:      false,
						Italic:    false,
						Family:    "Times New Roman",
						Size:      14,
						Color:     "777777",
						VertAlign: "subscript",
					},
				},
			},
			Width:  180,
			Height: 40,
		},
	)
	f.AddShape("Sheet1",
		&excelize.Shape{
			Cell:   "K19",
			Type:   "rightArrow",
			Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
			Fill:   excelize.Fill{Color: []string{"FFFFFF"}, Pattern: 1},
			Width:  80,
			Height: 40,
		},
	)
	f.AddShape("Sheet1",
		&excelize.Shape{
			Cell: "O19",
			Type: "rect",
			Line: excelize.ShapeLine{Color: "060270", Width: &lineWidth},
			Fill: excelize.Fill{Color: []string{"FFFFFF"}, Pattern: 1},
			Paragraph: []excelize.RichTextRun{
				{
					Text: "Process True",
					Font: &excelize.Font{
						Bold:      false,
						Italic:    false,
						Family:    "Times New Roman",
						Size:      14,
						Color:     "777777",
						VertAlign: "superscript",
					},
				},
			},
			Width:  180,
			Height: 40,
		},
	)
	if err := f.SaveAs("Book4.xlsx"); err != nil {
		fmt.Println(err)
	}
}
