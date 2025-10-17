package handler

import (
	"fmt"
	"go_excelize/internal/app/service"
	"math/rand"
	"net/http"
	"strings"

	"github.com/xuri/excelize/v2"
)

type ExcelHandler struct {
	service *service.ExcelService
}

func NewExcelHandler(s *service.ExcelService) *ExcelHandler {
	return &ExcelHandler{service: s}
}

func newFlowchartShape(cell, shapeType string) *excelize.Shape {
	lineWidth := 1.2
	return &excelize.Shape{
		Cell: cell,
		Type: shapeType,
		Line: excelize.ShapeLine{Color: "060270", Width: &lineWidth},
		Fill: excelize.Fill{Color: []string{"FFFFFF"}, Pattern: 1},
		Paragraph: []excelize.RichTextRun{
			{
				// Text: text, // Use the text parameter here
				Font: &excelize.Font{
					Bold:   false,
					Italic: false,
					Family: "Times New Roman",
					Size:   14,
					Color:  "777777",
				},
			},
		},
		Width:  80,
		Height: 18,
	}
}

func newArrowShape(cell, orientation string) *excelize.Shape {
	lineWidth := 1.2
	var arrowType = ""
	var width = 0
	var height = 0

	if orientation == "up" {
		arrowType = "upArrow"
		width = 40
		height = 80
	} else if orientation == "down" {
		arrowType = "downArrow"
		width = 40
		height = 80
	} else if orientation == "right" {
		arrowType = "rightArrow"
		width = 40
		height = 80
	} else if orientation == "left" {
		arrowType = "leftArrow"
		width = 40
		height = 80
	}

	return &excelize.Shape{
		Cell:   cell,
		Type:   arrowType,
		Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
		Fill:   excelize.Fill{Color: []string{"FFFFFF"}, Pattern: 1},
		Width:  uint(width),
		Height: uint(height),
	}
}

func convertLengthToPixel(length float64) float64 {
	return (length - 5.5) / 7
}

func (h *ExcelHandler) GenerateExcel(w http.ResponseWriter, r *http.Request) {
	shapesParam := r.URL.Query().Get("shapes")
	startingCell := r.URL.Query().Get("start")
	orderParam := r.URL.Query().Get("orders")
	// textsParam := r.URL.Query().Get("texts")

	// --- Set a default value if the parameter is not provided ---
	if shapesParam == "" || startingCell == "" || orderParam == "" {
		http.Error(w, "Please provide 'shapes' and 'texts' query parameters as comma-separated lists.", http.StatusBadRequest)
		return
	}

	shapeTypes := strings.Split(shapesParam, ",")
	orderFlows := strings.Split(orderParam, ",")
	// shapeTexts := strings.Split(textsParam, ",")

	// --- Validate the input ---
	if len(shapeTypes) != len(orderFlows) {
		http.Error(w, "The number of shapes must match the number of texts.", http.StatusBadRequest)
		return
	}

	file := excelize.NewFile()
	file.SetColWidth("Sheet1", startingCell, startingCell, convertLengthToPixel(80))
	defer func() {
		if err := file.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// We'll use this variable to keep track of where to place the next object.
	currentRow := 6
	shapeColumn := startingCell
	// arrowColumn := "H" // Arrows are slightly offset for a cleaner look

	for i, shapeType := range shapeTypes {
		if i > 0 {
			// arrowCell := fmt.Sprintf("%s%d", arrowColumn, currentRow)
			// arrow := newArrowShape(arrowCell, "down")
			// file.AddShape("Sheet1", arrow)
			// // Move down the sheet to make space for the arrow
			// currentRow += 6
		}
		// Place the main flowchart shape
		shapeCell := fmt.Sprintf("%s%d", shapeColumn, currentRow)
		// text := shapeTexts[i]
		shape := newFlowchartShape(shapeCell, shapeType)
		file.AddShape("Sheet1", shape)

		// Move down the sheet to make space for the shape for the next loop iteration
		file.SetRowHeight("Sheet1", currentRow, 20)
		currentRow += 1
	}

	filename := fmt.Sprintf("flowchart_%d.xlsx", rand.Intn(10000))

	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	w.Header().Set("Content-Disposition", "attachment; filename="+filename)
	if err := file.Write(w); err != nil {
		http.Error(w, "Failed to generate file", http.StatusInternalServerError)
		return
	}
}
