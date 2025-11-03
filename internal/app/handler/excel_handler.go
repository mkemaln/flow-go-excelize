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

func newArrowShape(cell, originPos, targetPos, orientation string, shapeWidth, shapeHeight, cellWidth, cellHeight float64, arrowLength int) (*excelize.Shape, *excelize.Shape) {
	// lineWidth := 0.25
	lineWidth := 2.0
	// line := &excelize.Shape{
	// 	Cell: cell,
	// 	Line: excelize.ShapeLine{Color: "060270", Width: &lineWidth},
	// 	Fill: excelize.Fill{Color: []string{"060270"}, Pattern: 1},
	// }

	// arrowHead := &excelize.Shape{
	// 	Line: excelize.ShapeLine{Color: "060270", Width: &lineWidth},
	// 	Fill: excelize.Fill{Color: []string{"060270"}, Pattern: 1},
	// }
	line := &excelize.Shape{
		Cell: cell,
		Line: excelize.ShapeLine{Color: "000000", Width: &lineWidth},
		Fill: excelize.Fill{Color: []string{"000000"}, Pattern: 1},
	}

	arrowHead := &excelize.Shape{
		Line: excelize.ShapeLine{Color: "000000", Width: &lineWidth},
		Fill: excelize.Fill{Color: []string{"000000"}, Pattern: 1},
	}

	// widthDiff is now SIGNED. e.g., -1 for left, +1 for right.
	widthDiff, heightDiff, err := getCellDifference(originPos, targetPos)
	if err != nil {
		fmt.Println("Error:", err)
	}

	colNum, rowNum, err := excelize.CellNameToCoordinates(cell)
	if err != nil {
		fmt.Println("Error:", err)
	}

	// Use absolute values for width/height calculations
	absWidthDiff := math.Abs(float64(widthDiff))
	absHeightDiff := math.Abs(float64(heightDiff))

	switch orientation {
	case "downConn":
		arrowX := cellWidth / 2
		arrowY := (cellHeight-shapeHeight)/2 + shapeHeight - (cellHeight-shapeHeight)/2 // - 5 is added because for some reason there were a down offset

		line.Type = "downArrow"
		line.Width = 1
		line.Height = uint(arrowLength)
		line.Format = excelize.GraphicOptions{
			OffsetX: int(arrowX),
			OffsetY: int(arrowY),
		}

		arrowHead = nil

	case "rightConn":
		arrowX := (cellWidth-shapeWidth)/2 + shapeWidth - (cellWidth-shapeWidth)/2
		arrowY := cellHeight / 2
		arrowWidth := (cellWidth-shapeWidth)/2 + cellWidth*(absWidthDiff-1) + cellWidth/2 + (cellWidth - shapeWidth) // add 1 padding
		arrowHeight := (cellHeight-shapeHeight)/2 + cellHeight*(absHeightDiff-1) + cellHeight/2

		line.Type = "rect"
		line.Cell = cell
		line.Width = uint(arrowWidth)
		line.Height = 1
		line.Format = excelize.GraphicOptions{
			OffsetX: int(arrowX),
			OffsetY: int(arrowY),
		}

		arrowHead.Type = "downArrow"
		arrowHead.Cell = fmt.Sprintf("%c%d", 'A'+(colNum+widthDiff-1), rowNum)
		arrowHead.Width = 1
		arrowHead.Height = uint(arrowHeight)
		arrowHead.Format = excelize.GraphicOptions{
			OffsetX: int(cellWidth) / 2,
			OffsetY: int(cellHeight) / 2,
		}
	case "leftConn":
		arrowX := cellWidth / 2
		arrowY := cellHeight / 2
		arrowWidth := (cellWidth-shapeWidth)/2 + cellWidth*(absWidthDiff-1) + cellWidth/2
		arrowHeight := (cellHeight-shapeHeight)/2 + cellHeight*(absHeightDiff-1) + cellHeight/2

		line.Type = "rect"
		// --- FIX: Anchor to target cell ---
		line.Cell = targetPos
		line.Width = uint(arrowWidth)
		line.Height = 1
		line.Format = excelize.GraphicOptions{
			OffsetX: int(arrowX),
			OffsetY: int(arrowY),
		}

		arrowHead.Type = "downArrow"
		// --- FIX: Anchor to target cell ---
		arrowHead.Cell = targetPos
		arrowHead.Width = 1
		arrowHead.Height = uint(arrowHeight)
		arrowHead.Format = excelize.GraphicOptions{
			OffsetX: int(cellWidth) / 2,
			OffsetY: int(cellHeight) / 2,
		}
	case "upperRightConn":
		arrowX := (cellWidth-shapeWidth)/2 + shapeWidth - 5
		arrowY := cellHeight / 2
		arrowWidth := (cellWidth-shapeWidth)/2 + cellWidth*(absWidthDiff-1) + cellWidth/2
		arrowHeight := (cellHeight-shapeHeight)/2 + cellHeight*(absHeightDiff-1) + cellHeight/2

		line.Type = "rect"
		line.Cell = originPos
		line.Width = uint(arrowWidth)
		line.Height = 1
		line.Format = excelize.GraphicOptions{
			OffsetX: int(arrowX),
			OffsetY: int(arrowY),
		}

		arrowHead.Type = "upArrow"
		arrowHead.Cell = fmt.Sprintf("%c%d", 'A'+(colNum+widthDiff-1), rowNum)
		arrowHead.Width = 1
		arrowHead.Height = uint(arrowHeight)
		arrowHead.Format = excelize.GraphicOptions{
			OffsetX: int(cellWidth) / 2,
			OffsetY: int(cellHeight) / 2,
		}
	case "upperLeftConn":
		arrowX := cellWidth / 2
		arrowY := cellHeight / 2
		arrowWidth := (cellWidth-shapeWidth)/2 + cellWidth*(absWidthDiff-1) + cellWidth/2
		arrowHeight := (cellHeight-shapeHeight)/2 + cellHeight*(absHeightDiff-1) + cellHeight/2

		line.Type = "rect"
		line.Cell = originPos
		line.Width = uint(arrowWidth)
		line.Height = 1
		line.Format = excelize.GraphicOptions{
			OffsetX: int(arrowX),
			OffsetY: int(arrowY),
		}

		arrowHead.Type = "upArrow"
		arrowHead.Cell = fmt.Sprintf("%c%d", 'A'+(colNum+widthDiff-1), rowNum)
		arrowHead.Width = 1
		arrowHead.Height = uint(arrowHeight)
		arrowHead.Format = excelize.GraphicOptions{
			OffsetX: int(cellWidth) / 2,
			OffsetY: int(cellHeight) / 2,
		}
	}
	return line, arrowHead
}

// --- REMOVED findCellByOrder function ---

// --- NEW HELPER FUNCTION ---
// Parses a branch query string (e.g., "1:2,4:5") into a map[originIndex]targetIndex
func parseBranchParam(param string) (map[int]int, error) {
	branches := make(map[int]int)
	if param == "" {
		return branches, nil
	}
	// param is "1:2,4:5"
	pairs := strings.Split(param, ",") // ["1:2", "4:5"]
	for _, pair := range pairs {
		parts := strings.Split(pair, ":") // ["1", "2"]
		if len(parts) != 2 {
			return nil, fmt.Errorf("invalid branch format: %s", pair)
		}
		originIndex, err1 := strconv.Atoi(parts[0])
		targetIndex, err2 := strconv.Atoi(parts[1]) // Now parses a target *index*
		if err1 != nil || err2 != nil {
			return nil, fmt.Errorf("invalid branch numbers: %s", pair)
		}
		branches[originIndex] = targetIndex
	}
	return branches, nil
}

func (h *ExcelHandler) GenerateExcel(w http.ResponseWriter, r *http.Request) {
	// --- Read new parameters ---
	shapesParam := r.URL.Query().Get("shapes")
	startCellParam := r.URL.Query().Get("start")
	orderParam := r.URL.Query().Get("orders")
	trueBranchesParam := r.URL.Query().Get("true_branches")
	falseBranchesParam := r.URL.Query().Get("false_branches")
	shapeWidthParam := r.URL.Query().Get("width")
	shapeHeightParam := r.URL.Query().Get("height")
	cellPadParam := r.URL.Query().Get("pad")
	gapParam := r.URL.Query().Get("gap")

	if shapesParam == "" || startCellParam == "" || orderParam == "" || shapeWidthParam == "" || shapeHeightParam == "" || gapParam == "" || cellPadParam == "" {
		http.Error(w, "Please provide all required parameters.", http.StatusBadRequest)
		return
	}

	shapeTypes := strings.Split(shapesParam, ",")
	orderFlows := strings.Split(orderParam, ",")
	shapeWidth, _ := strconv.Atoi(shapeWidthParam)
	shapeHeight, _ := strconv.Atoi(shapeHeightParam)
	verticalGap, _ := strconv.Atoi(gapParam)
	cellPadding, _ := strconv.Atoi(cellPadParam)

	// --- Parse new branch parameters (now index-to-index) ---
	trueBranches, errT := parseBranchParam(trueBranchesParam)
	if errT != nil {
		http.Error(w, fmt.Sprintf("Invalid 'true_branches' param: %v", errT), http.StatusBadRequest)
		return
	}
	falseBranches, errF := parseBranchParam(falseBranchesParam)
	if errF != nil {
		http.Error(w, fmt.Sprintf("Invalid 'false_branches' param: %v", errF), http.StatusBadRequest)
		return
	}

	if len(shapeTypes) != len(orderFlows) {
		http.Error(w, "The number of shapes and orders must all match.", http.StatusBadRequest)
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
	startColIndex := startColNum - 1
	startRow := startRowNum

	colRows := make(map[int]int)
	var prevColIndex int
	var prevRow int

	// --- Maps to store locations and branch logic ---
	shapeLocations := make(map[int]string) // Key: shape index (i), Value: "G6"
	// shapeOrders map is no longer needed for branching

	// --- FIRST LOOP: Place shapes and record information ---
	for i, shapeType := range shapeTypes {
		orderFlow := orderFlows[i]

		// --- Simplified 'orders' parsing (complex validation removed) ---
		order, err := strconv.Atoi(orderFlow)
		if err != nil {
			http.Error(w, fmt.Sprintf("Invalid order number for shape at index %d: %s", i, orderFlow), http.StatusBadRequest)
			return
		}

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

		cellWidth := float64(shapeWidth + cellPadding)
		cellHeight := float64(shapeHeight + cellPadding)
		file.SetColWidth("Sheet1", currentColName, currentColName, pixelsToCharUnits(cellWidth))
		file.SetRowHeight("Sheet1", currentRow, pixelsToPoints(cellHeight))

		shape := newFlowchartShape(currentShapeCell, shapeType, uint(shapeWidth), uint(shapeHeight), uint(cellPadding))
		file.AddShape("Sheet1", shape)

		// --- RECORD data for the second loop ---
		shapeLocations[i] = currentShapeCell
		// shapeOrders[i] = order // No longer needed for branching

		// Update state for the next loop
		colRows[currentColIndex] = currentRow + verticalGap
		prevColIndex = currentColIndex
		prevRow = colRows[currentColIndex]
	}

	// --- SECOND LOOP: Draw all arrows ---
	for i, shapeType := range shapeTypes {
		originCell := shapeLocations[i]
		cellWidth := float64(shapeWidth + cellPadding)
		cellHeight := float64(shapeHeight + cellPadding)

		isDecision := (shapeType == "flowChartDecision")
		hasTrueBranch := false
		hasFalseBranch := false

		// --- Logic for drawing "true" branch ---
		if targetIndex, exists := trueBranches[i]; exists {
			hasTrueBranch = true
			if targetCell, ok := shapeLocations[targetIndex]; ok {
				// We have an origin and a target cell, draw the arrow
				originCol, _, _ := excelize.CellNameToCoordinates(originCell)
				targetCol, _, _ := excelize.CellNameToCoordinates(targetCell)
				orientation := "downConn" // default
				arrowCell := originCell
				if targetCol > originCol {
					orientation = "rightConn"
				} else if targetCol < originCol {
					orientation = "leftConn"
					arrowCell = targetCell // Anchor left to target
				}

				line, arrowhead := newArrowShape(arrowCell, originCell, targetCell, orientation, float64(shapeWidth), float64(shapeHeight), cellWidth, cellHeight, cellPadding)
				file.AddShape("Sheet1", line)
				if arrowhead != nil {
					file.AddShape("Sheet1", arrowhead)
				}
			}
		}

		// --- Logic for drawing "false" branch ---
		if targetIndex, exists := falseBranches[i]; exists {
			hasFalseBranch = true
			if targetCell, ok := shapeLocations[targetIndex]; ok {
				// We have an origin and a target cell, draw the arrow
				originCol, _, _ := excelize.CellNameToCoordinates(originCell)
				targetCol, _, _ := excelize.CellNameToCoordinates(targetCell)
				orientation := "upperRightConn" // default
				arrowCell := originCell
				if targetCol > originCol {
					orientation = "upperRightConn"
				} else if targetCol < originCol {
					orientation = "upperLeftConn"
					arrowCell = targetCell // Anchor left to target
				}

				line, arrowhead := newArrowShape(arrowCell, originCell, targetCell, orientation, float64(shapeWidth), float64(shapeHeight), cellWidth, cellHeight, cellPadding)
				file.AddShape("Sheet1", line)
				if arrowhead != nil {
					file.AddShape("Sheet1", arrowhead)
				}
			}
		}

		// --- Logic for simple sequential connection ---
		// If this is NOT a decision, and it does NOT have a custom branch,
		// and it is NOT the last shape in the list...
		if !isDecision && !hasTrueBranch && !hasFalseBranch && i < len(shapeTypes)-1 {
			// Connect it to the next shape in the list (i+1)
			targetCell := shapeLocations[i+1]
			originCol, _, _ := excelize.CellNameToCoordinates(originCell)
			targetCol, _, _ := excelize.CellNameToCoordinates(targetCell)

			var orientation string
			var arrowCell string
			if originCol == targetCol {
				orientation = "downConn"
				arrowCell = originCell
			} else if targetCol > originCol {
				orientation = "rightConn"
				arrowCell = originCell
			} else {
				orientation = "leftConn"
				arrowCell = targetCell // Anchor to the target cell for left connections
			}

			line, arrowhead := newArrowShape(arrowCell, originCell, targetCell, orientation, float64(shapeWidth), float64(shapeHeight), cellWidth, cellHeight, cellPadding)
			file.AddShape("Sheet1", line)
			if arrowhead != nil {
				file.AddShape("Sheet1", arrowhead)
			}
		}
	}

	// ... (rest of your handler, writing the file) ...
	filename := fmt.Sprintf("flowchart_%d.xlsx", rand.Intn(10000))
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument/spreadsheetml.sheet")
	w.Header().Set("Content-Disposition", "attachment; filename="+filename)
	if err := file.Write(w); err != nil {
		http.Error(w, "Failed to generate file", http.StatusInternalServerError)
	}
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

// --- CRITICAL FIX: Removed math.Abs ---
// We need the signed difference to determine direction (left vs. right)
func getCellDifference(cell1, cell2 string) (width, height int, err error) {
	col1, row1, err := excelize.CellNameToCoordinates(cell1)
	if err != nil {
		return 0, 0, fmt.Errorf("invalid first cell name: %w", err)
	}

	col2, row2, err := excelize.CellNameToCoordinates(cell2)
	if err != nil {
		return 0, 0, fmt.Errorf("invalid second cell name: %w", err)
	}

	// Return the signed difference
	width = col2 - col1
	height = row2 - row1

	return width, height, nil
}
