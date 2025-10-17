package handler

import (
	"fmt"
	"go_excelize/internal/app/service"
	"math/rand"
	"net/http"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

type ExcelHandler struct {
	service *service.ExcelService
}

func NewExcelHandler(s *service.ExcelService) *ExcelHandler {
	return &ExcelHandler{service: s}
}

// Pass the text in, but not the cell dimensions.
func newFlowchartShape(cell, shapeType, text string, width, height uint) *excelize.Shape {
	lineWidth := 1.2
	return &excelize.Shape{
		Cell: cell,
		Type: shapeType,
		Line: excelize.ShapeLine{Color: "060270", Width: &lineWidth},
		Fill: excelize.Fill{Color: []string{"FFFFFF"}, Pattern: 1},
		Paragraph: []excelize.RichTextRun{
			{
				Text: text, // Use the text parameter here
				Font: &excelize.Font{
					Bold:   false,
					Italic: false,
					Family: "Times New Roman",
					Size:   14,
					Color:  "777777",
				},
			},
		},
		Width:  width,
		Height: height,
		// --- THIS IS THE CORRECT FIX ---
		// Use 'Format' and 'Positioning', not 'GraphicOptions' for this problem.
		Format: excelize.GraphicOptions{
			Positioning: "oneCell", // "Move but do not size with cells"
		},
	}
}

func newArrowShape(cell, orientation string) *excelize.Shape {
	lineWidth := 1.2
	shape := &excelize.Shape{
		Cell: cell,
		Line: excelize.ShapeLine{Color: "060270", Width: &lineWidth},
		Fill: excelize.Fill{Color: []string{"FFFFFF"}, Pattern: 1},
		// Add the fix here as well for consistency
		Format: excelize.GraphicOptions{
			Positioning: "oneCell",
		},
	}

	switch orientation {
	case "up":
		shape.Type = "upArrow"
		shape.Width = 40
		shape.Height = 80
	case "down":
		shape.Type = "downArrow"
		shape.Width = 40
		shape.Height = 80
	case "right":
		shape.Type = "rightArrow"
		shape.Width = 80
		shape.Height = 40
	case "left":
		shape.Type = "leftArrow"
		shape.Width = 80
		shape.Height = 40
	}
	return shape
}

// Pixels to points (for SetRowHeight)
func pixelsToPoints(pixels float64) float64 {
	if pixels == 0 {
		return 0
	}
	return pixels * 3.0 / 4.0 // Adjusted for better accuracy
}

// Pixels to character units (for SetColWidth)
func pixelsToCharUnits(pixels float64) float64 {
	if pixels <= 0 {
		return 0
	}
	// This is an approximation, Excel's calculation is complex
	return (pixels - 5) / 7
}

func (h *ExcelHandler) GenerateExcel(w http.ResponseWriter, r *http.Request) {
	// Re-added 'texts' as it's needed
	shapesParam := r.URL.Query().Get("shapes")
	startColumn := r.URL.Query().Get("start")
	orderParam := r.URL.Query().Get("orders")
	textsParam := r.URL.Query().Get("texts")
	shapeWidthParam := r.URL.Query().Get("width")
	shapeHeightParam := r.URL.Query().Get("height")
	gapParam := r.URL.Query().Get("gap")

	if shapesParam == "" || startColumn == "" || orderParam == "" || textsParam == "" || shapeWidthParam == "" || shapeHeightParam == "" || gapParam == "" {
		http.Error(w, "Please provide 'shapes', 'start', 'orders', 'texts', 'width', 'height', and 'gap' parameters.", http.StatusBadRequest)
		return
	}

	shapeTypes := strings.Split(shapesParam, ",")
	orderFlows := strings.Split(orderParam, ",")
	shapeTexts := strings.Split(textsParam, ",")

	shapeWidth, _ := strconv.Atoi(shapeWidthParam)
	shapeHeight, _ := strconv.Atoi(shapeHeightParam)
	verticalGap, _ := strconv.Atoi(gapParam)

	if len(shapeTypes) != len(orderFlows) || len(shapeTypes) != len(shapeTexts) {
		http.Error(w, "The number of shapes, orders, and texts must all match.", http.StatusBadRequest)
		return
	}

	file := excelize.NewFile()
	defer func() {
		if err := file.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	startColIndex := int(strings.ToUpper(startColumn)[0] - 'A')
	startRow := 6
	colRows := make(map[int]int)
	// var prevShapeCell string
	var prevColIndex int
	var prevRow int

	for i, shapeType := range shapeTypes {
		order, _ := strconv.Atoi(orderFlows[i])
		currentColIndex := startColIndex + (order - 1)
		var currentRow int

		if i == 0 {
			currentRow = startRow
		} else if currentColIndex == prevColIndex {
			currentRow = colRows[currentColIndex]
		} else {
			currentRow = prevRow
		}

		currentShapeCell := fmt.Sprintf("%c%d", 'A'+currentColIndex, currentRow)
		currentColName, _ := excelize.ColumnNumberToName(currentColIndex + 1)

		// Set cell dimensions before placing the shape
		// Cell width should be wider than the shape for good layout
		file.SetColWidth("Sheet1", currentColName, currentColName, pixelsToCharUnits(float64(shapeWidth+40)))
		file.SetRowHeight("Sheet1", currentRow, pixelsToPoints(float64(shapeHeight+20)))

		if i > 0 {
			// var orientation string
			// var arrowCell string
			// if currentColIndex > prevColIndex {
			// 	orientation = "right"
			// 	arrowCell = prevShapeCell
			// } else if currentColIndex < prevColIndex {
			// 	orientation = "left"
			// 	arrowCell = currentShapeCell
			// } else {
			// 	orientation = "down"
			// 	arrowCell = fmt.Sprintf("%c%d", 'A'+prevColIndex, currentRow-verticalGap/2)
			// }
			// arrow := newArrowShape(arrowCell, orientation)
			// file.AddShape("Sheet1", arrow)
		}

		shape := newFlowchartShape(currentShapeCell, shapeType, shapeTexts[i], uint(shapeWidth), uint(shapeHeight))
		file.AddShape("Sheet1", shape)

		// Update state for the next loop
		colRows[currentColIndex] = currentRow + verticalGap
		// prevShapeCell = currentShapeCell
		prevColIndex = currentColIndex
		// --- LOGIC FIX HERE ---
		// prevRow should be the current row, not the next available row.
		prevRow = colRows[currentColIndex]
	}

	filename := fmt.Sprintf("flowchart_%d.xlsx", rand.Intn(10000))
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	w.Header().Set("Content-Disposition", "attachment; filename="+filename)
	if err := file.Write(w); err != nil {
		http.Error(w, "Failed to generate file", http.StatusInternalServerError)
	}
}
