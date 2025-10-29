package main

import (
	"fmt"
	"log"

	"github.com/xuri/excelize/v2"
)

func main() {
	// Create test Excel file to demonstrate current arrow issues
	file := excelize.NewFile()
	defer file.Close()

	// Create shapes
	shapes := []struct {
		cell, shapeType string
	}{
		{"D4", "rect"},
		{"D6", "rect"},
		{"E6", "rect"},
		{"G6", "rect"},
		{"F6", "rect"},
	}

	lineWidth := 0.5

	for i, shape := range shapes {
		// Add main shape
		mainShape := &excelize.Shape{
			Cell:   shape.cell,
			Type:   shape.shapeType,
			Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
			Fill:   excelize.Fill{Color: []string{"FFFFFF"}, Pattern: 1},
			Width:  80,
			Height: 40,
		}
		file.AddShape("Sheet1", mainShape)

		// Try different arrow approaches
		if i < len(shapes)-1 {
			// Approach 1: Using rect as line (current problematic approach)
			rectLine := &excelize.Shape{
				Cell:   shape.cell,
				Type:   "rect", // This creates a rectangle, not a line
				Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
				Fill:   excelize.Fill{Color: []string{"060270"}, Pattern: 1},
				Width:  30, // This will appear as a rectangle, not a line
				Height: 1,
				Format: excelize.GraphicOptions{
					OffsetX: 85,
					OffsetY: 20,
				},
			}
			file.AddShape("Sheet1", rectLine)

			// Add arrow head
			arrowHead := &excelize.Shape{
				Type:   "rightArrow",
				Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
				Fill:   excelize.Fill{Color: []string{"060270"}, Pattern: 1},
				Width:  uint(lineWidth * 4),
				Height: uint(lineWidth * 3),
				Format: excelize.GraphicOptions{
					OffsetX: 115,
					OffsetY: 18,
				},
			}
			file.AddShape("Sheet1", arrowHead)
		}
	}

	// Save the file
	if err := file.WriteAsFile("arrow_test_current_issues.xlsx"); err != nil {
		log.Fatal(err)
	}

	fmt.Println("Test file created: arrow_test_current_issues.xlsx")
	fmt.Println("This demonstrates the current issues:")
	fmt.Println("1. Using 'rect' shapes creates rectangles, not lines")
	fmt.Println("2. Lines always appear as rectangles regardless of dimensions")
	fmt.Println("3. Limited arrow rotation capabilities")
}
