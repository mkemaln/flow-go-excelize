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
	lineWidth := 0.25
	line := &excelize.Shape{
		Cell: cell,
		Line: excelize.ShapeLine{Color: "060270", Width: &lineWidth},
		Fill: excelize.Fill{Color: []string{"060270"}, Pattern: 1},
	}

	arrowHead := &excelize.Shape{
		Line: excelize.ShapeLine{Color: "060270", Width: &lineWidth},
		Fill: excelize.Fill{Color: []string{"060270"}, Pattern: 1},
	}

	widthDiff, heightDiff, err := getCellDifference(originPos, targetPos)
	if err != nil {
		fmt.Println("Error:", err)
	}

	colNum, rowNum, err := excelize.CellNameToCoordinates(cell)
	if err != nil {
		fmt.Println("Error:", err)
	}

	switch orientation {
	case "downConn":
		arrowX := cellWidth / 2
		arrowY := (cellHeight-shapeHeight)/2 + shapeHeight - 5 // - 5 is added because for some reason there were a down offset

		line.Type = "downArrow"
		line.Width = 1
		line.Height = uint(arrowLength)
		line.Format = excelize.GraphicOptions{
			OffsetX: int(arrowX),
			OffsetY: int(arrowY),
		}

		arrowHead = nil

	case "rightConn":
		arrowX := (cellWidth-shapeWidth)/2 + shapeWidth - 5
		arrowY := cellHeight / 2
		arrowWidth := (cellWidth-shapeWidth)/2 + cellWidth*(float64(widthDiff)-1) + cellWidth/2
		arrowHeight := (cellHeight-shapeHeight)/2 + cellHeight*(float64(heightDiff)-1) + cellHeight/2

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
		arrowWidth := (cellWidth-shapeWidth)/2 + cellWidth*(float64(widthDiff)-1) + cellWidth/2
		arrowHeight := (cellHeight-shapeHeight)/2 + cellHeight*(float64(heightDiff)-1) + cellHeight/2

		line.Type = "rect"
		line.Cell = fmt.Sprintf("%c%d", 'A'+(colNum-widthDiff-1), rowNum)
		line.Width = uint(arrowWidth)
		line.Height = 1
		line.Format = excelize.GraphicOptions{
			OffsetX: int(arrowX),
			OffsetY: int(arrowY),
		}

		arrowHead.Type = "downArrow"
		arrowHead.Cell = fmt.Sprintf("%c%d", 'A'+(colNum-widthDiff-1), rowNum)
		arrowHead.Width = 1
		arrowHead.Height = uint(arrowHeight)
		arrowHead.Format = excelize.GraphicOptions{
			OffsetX: int(cellWidth) / 2,
			OffsetY: int(cellHeight) / 2,
		}
	case "uppperRightConn":
		arrowX := (cellWidth-shapeWidth)/2 + shapeWidth - 5
		arrowY := cellHeight / 2
		arrowWidth := (cellWidth-shapeWidth)/2 + cellWidth*(float64(widthDiff)-1) + cellWidth/2
		arrowHeight := (cellHeight-shapeHeight)/2 + cellHeight*(float64(heightDiff)-1) + cellHeight/2

		line.Type = "rect"
		line.Cell = cell
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
		line.Type = "a"
	}
	return line, arrowHead
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

	type BranchInfo struct {
		FalseBranch int
		TrueBranch  int
	}
	// This map will store the branch targets for our decision shapes
	decisionBranches := make(map[int]BranchInfo)

	for i, shapeType := range shapeTypes {
		var order int
		var falseBranchOrder int = -1 // -1 means "not applicable"
		var trueBranchOrder int = -1  // -1 means "not applicable"
		var err error                 // Declare a single error variable

		orderFlow := orderFlows[i] // Get the raw order string, e.g., "1" or "2:1,3"

		if shapeType == "flowChartDecision" {
			if !strings.Contains(orderFlow, ":") {
				msg := fmt.Sprintf("Shape at index %d is 'flowChartDecision' but is missing branch targets in 'orders' (e.g., '2:1,3').", i)
				http.Error(w, msg, http.StatusBadRequest)
				return
			}
			parts := strings.Split(orderFlow, ":")
			if len(parts) != 2 {
				http.Error(w, fmt.Sprintf("Invalid order format for decision at index %d: %s", i, orderFlow), http.StatusBadRequest)
				return
			}
			branchParts := strings.Split(parts[1], ",")
			if len(branchParts) != 2 {
				http.Error(w, fmt.Sprintf("Decision at index %d must have two branch targets (e.g., '2:1,3'). Found: %s", i, parts[1]), http.StatusBadRequest)
				return
			}

			// --- 2. ASSIGN values using = (not :=) ---
			var orderNum, fBranch, tBranch int
			var err1, err2, err3 error

			orderNum, err1 = strconv.Atoi(parts[0])
			fBranch, err2 = strconv.Atoi(branchParts[0])
			tBranch, err3 = strconv.Atoi(branchParts[1])

			if err1 != nil || err2 != nil || err3 != nil {
				http.Error(w, fmt.Sprintf("Invalid number in order/branch targets for decision at index %d: %s", i, orderFlow), http.StatusBadRequest)
				return
			}
			order = orderNum
			falseBranchOrder = fBranch
			trueBranchOrder = tBranch

			// Store the branch info in our map
			decisionBranches[i] = BranchInfo{
				FalseBranch: falseBranchOrder,
				TrueBranch:  trueBranchOrder,
			}

		} else {
			if strings.Contains(orderFlow, ":") {
				msg := fmt.Sprintf("Shape at index %d (%s) is not a decision, but 'orders' has branch targets: %s", i, shapeType, orderFlow)
				http.Error(w, msg, http.StatusBadRequest)
				return
			}

			// --- 3. ASSIGN values using = ---
			order, err = strconv.Atoi(orderFlow)
			if err != nil {
				http.Error(w, fmt.Sprintf("Invalid order number for shape at index %d: %s", i, orderFlow), http.StatusBadRequest)
				return
			}
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

		// Set cell dimensions before placing the shape
		// Cell width should be wider than the shape for good layout
		cellWidth := float64(shapeWidth + cellPadding)
		cellHeight := float64(shapeHeight + cellPadding)
		file.SetColWidth("Sheet1", currentColName, currentColName, pixelsToCharUnits(cellWidth))
		file.SetRowHeight("Sheet1", currentRow, pixelsToPoints(cellHeight))

		if i > 0 {
			var orientation string
			var arrowCell string
			if currentColIndex == prevColIndex {
				orientation = "downConn"
				arrowCell = prevShapeCell
			} else if currentColIndex > prevColIndex {
				if shapeType == "flowChartDecision" {
					orientation = "uppperRightConn"
					arrowCell = prevShapeCell
				} else {
					orientation = "rightConn"
					arrowCell = prevShapeCell
				}
			} else if currentColIndex < prevColIndex {
				orientation = "leftConn"
				arrowCell = prevShapeCell
			}

			line, arrowhead := newArrowShape(arrowCell, prevShapeCell, currentShapeCell, orientation, float64(shapeWidth), float64(shapeHeight), cellWidth, cellHeight, cellPadding)

			if shapeType == "flowChartDecision" {
				prevShapeCell = falseBranchOrder
				line, arrowhead := newArrowShape(arrowCell, prevShapeCell, currentShapeCell, orientation, float64(shapeWidth), float64(shapeHeight), cellWidth, cellHeight, cellPadding)
				file.AddShape("Sheet1", line)

				if arrowhead != nil {
					file.AddShape("Sheet1", arrowhead)
				}
			}

			file.AddShape("Sheet1", line)

			if arrowhead != nil {
				file.AddShape("Sheet1", arrowhead)
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

	filename := fmt.Sprintf("flowchart_%d.xlsx", rand.Intn(10000))
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
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
