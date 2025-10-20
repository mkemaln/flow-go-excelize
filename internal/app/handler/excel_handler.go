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

type debugLog struct {
	Iteration int
	Variable  string
	Value     string
}

type ExcelHandler struct {
	service *service.ExcelService
}

func NewExcelHandler(s *service.ExcelService) *ExcelHandler {
	return &ExcelHandler{service: s}
}

// Pass the text in, but not the cell dimensions.
func newFlowchartShape(cell, shapeType string, width, height, cellPadding uint) *excelize.Shape {
	lineWidth := 1.2
	return &excelize.Shape{
		Cell: cell,
		Type: shapeType,
		Line: excelize.ShapeLine{Color: "060270", Width: &lineWidth},
		Fill: excelize.Fill{Color: []string{"FFFFFF"}, Pattern: 1},
		Paragraph: []excelize.RichTextRun{
			{
				// Text: fmt.Sprintf("%d", cellPadding),
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
		Format: excelize.GraphicOptions{
			Positioning: "oneCell", // "Move but do not size with cells"
			OffsetX:     int(cellPadding),
			OffsetY:     int(cellPadding),
		},
	}
}

func newArrowShape(cell, originPos, targetPos, statusPos, orientation string, shapeWidth, shapeHeight, cellWidth, cellHeight float64, arrowLength int) *excelize.Shape {
	lineWidth := 1.2
	shape := &excelize.Shape{
		Cell: cell,
		Line: excelize.ShapeLine{Color: "060270", Width: &lineWidth},
		Paragraph: []excelize.RichTextRun{
			{
				Text: statusPos,
				Font: &excelize.Font{
					Bold:   false,
					Italic: false,
					Family: "Times New Roman",
					Size:   5,
					Color:  "777777",
				},
			},
		},
	}

	widthDiff, heightDiff, err := getCellDifference(originPos, targetPos)
	if err != nil {
		fmt.Println("Error:", err)
	}

	switch orientation {
	case "same":
		arrowX := cellWidth / 2
		arrowY := (cellHeight-shapeHeight)/2 + shapeHeight

		shape.Type = "line"
		shape.Width = 2
		shape.Height = uint(arrowLength)
		shape.Format = excelize.GraphicOptions{
			OffsetX: int(arrowX),
			OffsetY: int(arrowY),
		}
	case "differ":
		arrowX := (cellWidth-shapeWidth)/2 + shapeWidth
		arrowY := cellHeight / 2
		arrowWidth := (cellWidth-shapeWidth)/2 + cellWidth*(float64(widthDiff)-1) + cellWidth/2
		arrowHeight := (cellHeight-shapeHeight)/2 + cellHeight*(float64(heightDiff)-1) + cellHeight/2

		shape.Type = "bentConnector2"
		shape.Width = uint(arrowWidth)
		shape.Height = uint(arrowHeight)
		shape.Format = excelize.GraphicOptions{
			OffsetX: int(arrowX),
			OffsetY: int(arrowY),
		}
	}
	return shape
}

func getCellDifference(cell1, cell2 string) (width, height int, err error) {
	// Use excelize's helper to convert cell names like "G7" into coordinates.
	// This robustly handles columns like "A", "Z", "AA", etc.
	col1, row1, err := excelize.CellNameToCoordinates(cell1)
	if err != nil {
		return 0, 0, fmt.Errorf("invalid first cell name: %w", err)
	}

	col2, row2, err := excelize.CellNameToCoordinates(cell2)
	if err != nil {
		return 0, 0, fmt.Errorf("invalid second cell name: %w", err)
	}

	// Calculate the difference. We use math.Abs to ensure the distance is always positive.
	// For example, the distance from A to G is the same as G to A.
	width = int(math.Abs(float64(col2 - col1)))
	height = int(math.Abs(float64(row2 - row1)))

	return width, height, nil
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
	shapesParam := r.URL.Query().Get("shapes")
	startCellParam := r.URL.Query().Get("start")
	orderParam := r.URL.Query().Get("orders")
	shapeWidthParam := r.URL.Query().Get("width")
	shapeHeightParam := r.URL.Query().Get("height")
	cellPadParam := r.URL.Query().Get("pad")
	gapParam := r.URL.Query().Get("gap")

	if shapesParam == "" || startCellParam == "" || orderParam == "" || shapeWidthParam == "" || shapeHeightParam == "" || gapParam == "" || cellPadParam == "" {
		http.Error(w, "Please provide 'shapes', 'start', 'orders', 'texts', 'width', 'height', and 'gap' parameters.", http.StatusBadRequest)
		return
	}

	shapeTypes := strings.Split(shapesParam, ",")
	orderFlows := strings.Split(orderParam, ",")

	shapeWidth, _ := strconv.Atoi(shapeWidthParam)
	shapeHeight, _ := strconv.Atoi(shapeHeightParam)
	verticalGap, _ := strconv.Atoi(gapParam)
	cellPadding, _ := strconv.Atoi(cellPadParam)

	if len(shapeTypes) != len(orderFlows) {
		http.Error(w, "The number of shapes, orders, and texts must all match.", http.StatusBadRequest)
		return
	}

	file := excelize.NewFile()
	defer func() {
		if err := file.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	startColNum, startRowNum, err := excelize.CellNameToCoordinates(startCellParam)
	if err != nil {
		http.Error(w, "Invalid 'start' parameter. Must be a valid cell reference (e.g., 'G6', 'AA1').", http.StatusBadRequest)
		return
	}
	// Convert excelize's 1-based column to our 0-based index
	startColIndex := startColNum - 1
	startRow := startRowNum

	colRows := make(map[int]int)
	var prevShapeCell string
	var prevColIndex int
	var prevRow int

	var logs []debugLog

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
		cellWidth := float64(shapeWidth + cellPadding)
		cellHeight := float64(shapeHeight + cellPadding)
		file.SetColWidth("Sheet1", currentColName, currentColName, pixelsToCharUnits(cellWidth))
		file.SetRowHeight("Sheet1", currentRow, pixelsToPoints(cellHeight))

		// === LOG ===
		logs = append(logs, debugLog{i, "--- Iteration Start ---", "---"})
		logs = append(logs, debugLog{i, "Shape Type", shapeType})
		logs = append(logs, debugLog{i, "Order", orderFlows[i]})
		logs = append(logs, debugLog{i, "Prev Cell", prevShapeCell})
		logs = append(logs, debugLog{i, "Current Cell", currentShapeCell})
		logs = append(logs, debugLog{i, "Cell Width", fmt.Sprintf("%.2f", cellWidth)})
		logs = append(logs, debugLog{i, "Cell Height", fmt.Sprintf("%.2f", cellHeight)})

		if i > 0 {
			var orientation string
			var arrowCell string
			if currentColIndex == prevColIndex {
				orientation = "same"
				arrowCell = fmt.Sprintf("%c%d", 'A'+prevColIndex, (currentRow-verticalGap/2)-1)
			} else {
				orientation = "differ"
				arrowCell = prevShapeCell
			}
			// orientation = "same"
			statusPos := fmt.Sprintf("pv%s,cr%s,ar%s", prevShapeCell, currentShapeCell, arrowCell)
			// arrowCell = fmt.Sprintf("%c%d", 'A'+prevColIndex, currentRow-verticalGap/2)
			arrow := newArrowShape(arrowCell, prevShapeCell, currentShapeCell, statusPos, orientation, float64(shapeWidth), float64(shapeHeight), cellWidth, cellHeight, cellPadding)
			file.AddShape("Sheet1", arrow)

			// === LOG ===
			// --- 4. Add logs for the arrow calculations ---
			logs = append(logs, debugLog{i, "Arrow Orientation", orientation})
			logs = append(logs, debugLog{i, "Arrow Anchor Cell", arrowCell})
			// Re-calculating for logging purposes to not change newArrowShape function
			widthDiff, heightDiff, _ := getCellDifference(prevShapeCell, currentShapeCell)
			if orientation == "differ" {
				aw := (cellWidth-float64(shapeWidth))/2 + cellWidth*(float64(widthDiff)-1) + cellWidth/2
				ah := (cellHeight-float64(shapeHeight))/2 + cellHeight*(float64(heightDiff)-1) + cellHeight/2
				logs = append(logs, debugLog{i, "Arrow Width (calc)", fmt.Sprintf("%.2f", aw)})
				logs = append(logs, debugLog{i, "Arrow Height (calc)", fmt.Sprintf("%.2f", ah)})
			} else {
				logs = append(logs, debugLog{i, "Arrow Height (calc)", strconv.Itoa(cellPadding)})
			}
		}

		shape := newFlowchartShape(currentShapeCell, shapeType, uint(shapeWidth), uint(shapeHeight), uint(cellPadding))
		file.AddShape("Sheet1", shape)

		// Update state for the next loop
		prevShapeCell = currentShapeCell
		colRows[currentColIndex] = currentRow + verticalGap
		prevColIndex = currentColIndex
		prevRow = colRows[currentColIndex]
	}

	// LOG RENDER
	maxRow := 0
	for _, r := range colRows {
		if r > maxRow {
			maxRow = r
		}
	}
	// Start writing logs 5 rows below the last element
	debugStartRow := maxRow + 5
	// Set headers for the debug table
	file.SetCellValue("Sheet1", fmt.Sprintf("A%d", debugStartRow), "Debug Log Output")
	file.SetCellValue("Sheet1", fmt.Sprintf("A%d", debugStartRow+1), "Iteration")
	file.SetCellValue("Sheet1", fmt.Sprintf("B%d", debugStartRow+1), "Variable")
	file.SetCellValue("Sheet1", fmt.Sprintf("C%d", debugStartRow+1), "Value")

	// Write each log entry
	for i, logEntry := range logs {
		row := debugStartRow + 2 + i
		file.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), logEntry.Iteration)
		file.SetCellValue("Sheet1", fmt.Sprintf("B%d", row), logEntry.Variable)
		file.SetCellValue("Sheet1", fmt.Sprintf("C%d", row), logEntry.Value)
	}
	// END LOG RENDER

	filename := fmt.Sprintf("flowchart_%d.xlsx", rand.Intn(10000))
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	w.Header().Set("Content-Disposition", "attachment; filename="+filename)
	if err := file.Write(w); err != nil {
		http.Error(w, "Failed to generate file", http.StatusInternalServerError)
	}
}
