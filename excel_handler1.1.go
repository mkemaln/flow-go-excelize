package handler

import (
	"fmt"
	"go_excelize/internal/app/service"
	"math"
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

func newFlowchartShape(cell, shapeType string, width, height, colWidth, rowHeight int) *excelize.Shape {
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
		Width:  uint(width),
		Height: uint(height),
		Format: excelize.GraphicOptions{
			OffsetX:     (colWidth - width) / 2,
			OffsetY:     (rowHeight - height) / 2,
			ScaleX:      0.5,
			ScaleY:      0.5,
			Positioning: "oneCell",
		},
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

// Pixels to character units (for SetColWidth)
func pixelsToCharUnits(pixels float64) float64 {
	if pixels <= 0 {
		return 0
	}
	return (pixels - 5.5) / 7.0
}

// Character units to pixels (for calculations)
func charUnitsToPixels(charUnits float64) float64 {
	if charUnits == 0 {
		return 0
	}
	if charUnits < 1 {
		return (charUnits * 12) + 0.5
	}
	return (charUnits * 7) + 5.5
}

// Points to pixels (for row height)
func pointsToPixels(points float64) float64 {
	if points == 0 {
		return 0
	}
	return math.Ceil(4.0 / 3.4 * points)
}

// Pixels to points (for SetRowHeight)
func pixelsToPoints(pixels float64) float64 {
	if pixels == 0 {
		return 0
	}
	return pixels * 3.4 / 4.0
}

func (h *ExcelHandler) GenerateExcel(w http.ResponseWriter, r *http.Request) {
	shapesParam := r.URL.Query().Get("shapes")
	startColumn := r.URL.Query().Get("start")
	orderParam := r.URL.Query().Get("orders")
	colParam := r.URL.Query().Get("width")
	rowParam := r.URL.Query().Get("height")

	// --- Set a default value if the parameter is not provided ---
	if shapesParam == "" || startColumn == "" || orderParam == "" {
		http.Error(w, "Please provide 'shapes' and 'texts' query parameters as comma-separated lists.", http.StatusBadRequest)
		return
	}

	shapeTypes := strings.Split(shapesParam, ",")
	orderFlows := strings.Split(orderParam, ",")

	colWidth, err := strconv.Atoi(colParam)
	if err != nil {
		http.Error(w, "Invalid 'colWidth' parameter. Must be numbers.", http.StatusBadRequest)
		return
	}

	rowHeight, err := strconv.Atoi(rowParam)
	if err != nil {
		http.Error(w, "Invalid 'rowHeight' parameter. Must be numbers.", http.StatusBadRequest)
		return
	}

	// --- Validate the input ---
	if len(shapeTypes) != len(orderFlows) {
		http.Error(w, "The number of shapes must match the number of orders.", http.StatusBadRequest)
		return
	}

	file := excelize.NewFile()
	defer func() {
		if err := file.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// --- 1. Setup for Grid Calculation ---
	startColIndex := int(strings.ToUpper(startColumn)[0] - 'A')
	startRow := 6
	verticalSpacing := 1 // How many rows to jump for a vertical connection

	// This map tracks the next available row for each column index
	colRows := make(map[int]int)

	// var prevShapeCell string
	var prevColIndex int
	var prevRow int

	for i, shapeType := range shapeTypes {
		order, err := strconv.Atoi(orderFlows[i])
		if err != nil {
			http.Error(w, "Invalid 'orders' parameter. Must be numbers.", http.StatusBadRequest)
			return
		}

		// Calculate current shape's column and row
		currentColIndex := startColIndex + (order - 1)
		var currentRow int

		// --- 3. THE CORRECTED LOGIC ---
		if i == 0 {
			// This is the very first shape
			currentRow = startRow
		} else if currentColIndex == prevColIndex {
			// If we are in the SAME column, get the next available row from our map
			currentRow = colRows[currentColIndex]
		} else {
			// If we are in a NEW column, base our row on the PREVIOUS shape's row
			currentRow = prevRow
		}

		currentShapeCell := fmt.Sprintf("%c%d", 'A'+currentColIndex, currentRow)
		currentCol := fmt.Sprintf("%c", 'A'+currentColIndex)
		// currentCol := strconv.Itoa('A' + currentColIndex)

		if i > 0 {
			// var orientation string
			// var arrowCell string

			// if currentColIndex > prevColIndex {
			// 	orientation = "right"
			// 	arrowCell = prevShapeCell // Start arrow from the previous shape's cell
			// } else if currentColIndex < prevColIndex {
			// 	orientation = "left"
			// 	arrowCell = currentShapeCell // Start arrow from the current shape's cell for better alignment
			// } else {
			// 	orientation = "down"
			// 	// Place arrow between the two shapes vertically
			// 	arrowCell = fmt.Sprintf("%c%d", 'A'+prevColIndex, currentRow-verticalSpacing/2)
			// }
			// arrow := newArrowShape(arrowCell, orientation)
			// file.AddShape("Sheet1", arrow)
		}

		// Place the main flowchart shape
		shape := newFlowchartShape(currentShapeCell, shapeType, 80, 20, colWidth, rowHeight)
		file.AddShape("Sheet1", shape)

		file.SetRowHeight("Sheet1", currentRow, pixelsToPoints(float64(rowHeight)))
		file.SetColWidth("Sheet1", currentCol, currentCol, pixelsToCharUnits(float64(colWidth)))

		// Update state for the next loop
		colRows[currentColIndex] = currentRow + verticalSpacing
		// Store the current position for the next iteration to reference
		// prevShapeCell = currentShapeCell
		prevColIndex = currentColIndex
		prevRow = colRows[currentColIndex]
	}

	filename := fmt.Sprintf("flowchart_%d.xlsx", rand.Intn(10000))

	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	w.Header().Set("Content-Disposition", "attachment; filename="+filename)
	if err := file.Write(w); err != nil {
		http.Error(w, "Failed to generate file", http.StatusInternalServerError)
		return
	}
}
