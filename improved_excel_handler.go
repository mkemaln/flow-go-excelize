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

// Enhanced arrow creation using straight connectors instead of rect shapes
func createImprovedArrow(cell, originPos, targetPos, orientation string, shapeWidth, shapeHeight, cellWidth, cellHeight float64, arrowLength int) ([]*excelize.Shape, error) {
	var shapes []*excelize.Shape
	lineWidth := 0.5

	widthDiff, heightDiff, err := getCellDifference(originPos, targetPos)
	if err != nil {
		return nil, err
	}

	// Handle basic orientations with improved straight connectors
	switch orientation {
	case "downConn":
		// Vertical connector using downArrow
		downArrow := &excelize.Shape{
			Cell:   cell,
			Type:   "downArrow",
			Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
			Fill:   excelize.Fill{Color: []string{"060270"}, Pattern: 1},
			Width:  uint(lineWidth * 2),
			Height: uint(arrowLength),
			Format: excelize.GraphicOptions{
				OffsetX: int(cellWidth/2 - lineWidth),
				OffsetY: int((cellHeight-shapeHeight)/2 + shapeHeight),
			},
		}
		shapes = append(shapes, downArrow)

	case "rightConn":
		// Horizontal connector using straightConnector1
		if widthDiff > 0 {
			horizontalWidth := float64(widthDiff) * cellWidth
			horizontalLine := &excelize.Shape{
				Cell:   cell,
				Type:   "straightConnector1",
				Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
				Fill:   excelize.Fill{Color: []string{"060270"}, Pattern: 1},
				Width:  uint(horizontalWidth),
				Height: uint(lineWidth),
				Format: excelize.GraphicOptions{
					OffsetX: int((cellWidth-shapeWidth)/2 + shapeWidth),
					OffsetY: int(cellHeight/2 - lineWidth/2),
				},
			}
			shapes = append(shapes, horizontalLine)

			// Arrow head pointing right
			arrowHead := &excelize.Shape{
				Type:   "rightArrow",
				Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
				Fill:   excelize.Fill{Color: []string{"060270"}, Pattern: 1},
				Width:  uint(lineWidth * 4),
				Height: uint(lineWidth * 3),
				Format: excelize.GraphicOptions{
					OffsetX: int((cellWidth-shapeWidth)/2 + shapeWidth + horizontalWidth - 2),
					OffsetY: int(cellHeight/2 - lineWidth*1.5),
				},
			}
			shapes = append(shapes, arrowHead)
		}

	case "leftConn":
		// Horizontal connector using straightConnector1 (left)
		if widthDiff < 0 {
			horizontalWidth := float64(-widthDiff) * cellWidth
			horizontalLine := &excelize.Shape{
				Cell:   targetPos,
				Type:   "straightConnector1",
				Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
				Fill:   excelize.Fill{Color: []string{"060270"}, Pattern: 1},
				Width:  uint(horizontalWidth),
				Height: uint(lineWidth),
				Format: excelize.GraphicOptions{
					OffsetX: int(0),
					OffsetY: int(cellHeight/2 - lineWidth/2),
				},
			}
			shapes = append(shapes, horizontalLine)

			// Arrow head pointing left
			arrowHead := &excelize.Shape{
				Cell:   targetPos,
				Type:   "leftArrow",
				Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
				Fill:   excelize.Fill{Color: []string{"060270"}, Pattern: 1},
				Width:  uint(lineWidth * 4),
				Height: uint(lineWidth * 3),
				Format: excelize.GraphicOptions{
					OffsetX: int(-lineWidth * 2),
					OffsetY: int(cellHeight/2 - lineWidth*1.5),
				},
			}
			shapes = append(shapes, arrowHead)
		}
	}

	return shapes, nil
}

// Bent connector creation using L-shaped segments
func createBentConnector(originCell, targetCell string, originWidth, originHeight, targetWidth, targetHeight, cellWidth, cellHeight float64, direction string) ([]*excelize.Shape, error) {
	var shapes []*excelize.Shape
	lineWidth := 0.5

	// Parse cell coordinates
	originCol, originRow, err := excelize.CellNameToCoordinates(originCell)
	if err != nil {
		return nil, err
	}
	targetCol, targetRow, err := excelize.CellNameToCoordinates(targetCell)
	if err != nil {
		return nil, err
	}

	// Calculate the actual pixel positions
	originX := float64(originCol-1) * cellWidth
	originY := float64(originRow-1) * cellHeight
	targetX := float64(targetCol-1) * cellWidth
	targetY := float64(targetRow-1) * cellHeight

	// Calculate center points of shapes
	originCenterX := originX + originWidth/2
	originCenterY := originY + originHeight/2
	targetCenterX := targetX + targetWidth/2
	targetCenterY := targetY + targetHeight/2

	switch direction {
	case "rightDown":
		// L-shape: horizontal then vertical down
		horizontalEndX := targetCenterX - targetWidth/2
		horizontalEndY := originCenterY

		// First segment: horizontal line from origin to bend point
		if horizontalEndX > originCenterX {
			seg1Width := horizontalEndX - originCenterX
			seg1X := originCenterX + 5 // Small offset from shape edge
			seg1Y := originCenterY - (lineWidth / 2)

			shapes = append(shapes, &excelize.Shape{
				Type:   "straightConnector1",
				Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
				Fill:   excelize.Fill{Color: []string{"060270"}, Pattern: 1},
				Width:  uint(seg1Width),
				Height: uint(lineWidth),
				Format: excelize.GraphicOptions{
					OffsetX: int(seg1X),
					OffsetY: int(seg1Y),
				},
			})
		}

		// Second segment: vertical line from bend point to target
		if targetCenterY > originCenterY {
			seg2Height := targetCenterY - originCenterY
			seg2X := horizontalEndX - (lineWidth / 2)
			seg2Y := originCenterY + 5 // Small offset from horizontal line

			shapes = append(shapes, &excelize.Shape{
				Type:   "straightConnector1",
				Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
				Fill:   excelize.Fill{Color: []string{"060270"}, Pattern: 1},
				Width:  uint(lineWidth),
				Height: uint(seg2Height),
				Format: excelize.GraphicOptions{
					OffsetX: int(seg2X),
					OffsetY: int(seg2Y),
				},
			})
		}

		// Arrow head at target
		if targetCenterY > originCenterY {
			shapes = append(shapes, &excelize.Shape{
				Type:   "downArrow",
				Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
				Fill:   excelize.Fill{Color: []string{"060270"}, Pattern: 1},
				Width:  uint(lineWidth * 3),
				Height: uint(lineWidth * 6),
				Format: excelize.GraphicOptions{
					OffsetX: int(targetCenterX - (lineWidth * 1.5)),
					OffsetY: int(targetCenterY - targetHeight/2 - 10),
				},
			})
		}

	case "rightUp":
		// L-shape: horizontal then vertical up
		horizontalEndX := targetCenterX - targetWidth/2
		horizontalEndY := originCenterY

		// First segment: horizontal line from origin to bend point
		if horizontalEndX > originCenterX {
			seg1Width := horizontalEndX - originCenterX
			seg1X := originCenterX + 5
			seg1Y := originCenterY - (lineWidth / 2)

			shapes = append(shapes, &excelize.Shape{
				Type:   "straightConnector1",
				Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
				Fill:   excelize.Fill{Color: []string{"060270"}, Pattern: 1},
				Width:  uint(seg1Width),
				Height: uint(lineWidth),
				Format: excelize.GraphicOptions{
					OffsetX: int(seg1X),
					OffsetY: int(seg1Y),
				},
			})
		}

		// Second segment: vertical line from bend point to target (upwards)
		if targetCenterY < originCenterY {
			seg2Height := originCenterY - targetCenterY
			seg2X := horizontalEndX - (lineWidth / 2)
			seg2Y := targetCenterY + targetHeight/2 - 5

			shapes = append(shapes, &excelize.Shape{
				Type:   "straightConnector1",
				Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
				Fill:   excelize.Fill{Color: []string{"060270"}, Pattern: 1},
				Width:  uint(lineWidth),
				Height: uint(seg2Height),
				Format: excelize.GraphicOptions{
					OffsetX: int(seg2X),
					OffsetY: int(seg2Y),
				},
			})
		}

		// Arrow head at target (pointing up)
		if targetCenterY < originCenterY {
			shapes = append(shapes, &excelize.Shape{
				Type:   "upArrow",
				Line:   excelize.ShapeLine{Color: "060270", Width: &lineWidth},
				Fill:   excelize.Fill{Color: []string{"060270"}, Pattern: 1},
				Width:  uint(lineWidth * 3),
				Height: uint(lineWidth * 6),
				Format: excelize.GraphicOptions{
					OffsetX: int(targetCenterX - (lineWidth * 1.5)),
					OffsetY: int(targetCenterY + targetHeight/2 + 4),
				},
			})
		}
	}

	return shapes, nil
}

// Helper function to determine connector direction based on cell positions
func determineConnectorDirection(originCol, originRow, targetCol, targetRow int) string {
	if originCol == targetCol {
		if targetRow > originRow {
			return "downConn"
		} else {
			return "upConn" // Assuming upConn support exists
		}
	} else if targetCol > originCol {
		if targetRow > originRow {
			return "rightDown"
		} else {
			return "rightUp"
		}
	} else { // targetCol < originCol
		if targetRow > originRow {
			return "leftDown"
		} else {
			return "leftUp"
		}
	}
}

// Helper function to find cell by order
func findCellByOrder(targetOrder int, shapeOrders map[int]int, shapeLocations map[int]string) string {
	for i, order := range shapeOrders {
		if order == targetOrder {
			return shapeLocations[i]
		}
	}
	return ""
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
		http.Error(w, "Please provide all required parameters.", http.StatusBadRequest)
		return
	}

	shapeTypes := strings.Split(shapesParam, ",")
	orderFlows := strings.Split(orderParam, ",")
	shapeWidth, _ := strconv.Atoi(shapeWidthParam)
	shapeHeight, _ := strconv.Atoi(shapeHeightParam)
	verticalGap, _ := strconv.Atoi(gapParam)
	cellPadding, _ := strconv.Atoi(cellPadParam)

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

	// Maps to store locations and branch logic
	shapeLocations := make(map[int]string)
	shapeOrders := make(map[int]int)

	type BranchInfo struct {
		FalseTargetOrder int
		TrueTargetOrder  int
	}
	decisionBranches := make(map[int]BranchInfo)

	// FIRST LOOP: Place shapes and record information
	for i, shapeType := range shapeTypes {
		var order int
		var falseBranchOrder int = -1
		var trueBranchOrder int = -1
		var err error

		orderFlow := orderFlows[i]

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

			decisionBranches[i] = BranchInfo{
				FalseTargetOrder: falseBranchOrder,
				TrueTargetOrder:  trueBranchOrder,
			}

		} else {
			if strings.Contains(orderFlow, ":") {
				msg := fmt.Sprintf("Shape at index %d (%s) is not a decision, but 'orders' has branch targets: %s", i, shapeType, orderFlow)
				http.Error(w, msg, http.StatusBadRequest)
				return
			}
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

		cellWidth := float64(shapeWidth + cellPadding)
		cellHeight := float64(shapeHeight + cellPadding)
		file.SetColWidth("Sheet1", currentColName, currentColName, pixelsToCharUnits(cellWidth))
		file.SetRowHeight("Sheet1", currentRow, pixelsToPoints(cellHeight))

		shape := newFlowchartShape(currentShapeCell, shapeType, uint(shapeWidth), uint(shapeHeight), uint(cellPadding))
		file.AddShape("Sheet1", shape)

		shapeLocations[i] = currentShapeCell
		shapeOrders[i] = order

		colRows[currentColIndex] = currentRow + verticalGap
		prevColIndex = currentColIndex
		prevRow = colRows[currentColIndex]
	}

	// SECOND LOOP: Draw all arrows using improved connector functions
	for i, shapeType := range shapeTypes {
		originCell := shapeLocations[i]
		cellWidth := float64(shapeWidth + cellPadding)
		cellHeight := float64(shapeHeight + cellPadding)

		if shapeType == "flowChartDecision" {
			branchInfo := decisionBranches[i]

			// FALSE branch arrow
			targetCellFalse := findCellByOrder(branchInfo.FalseTargetOrder, shapeOrders, shapeLocations)
			if targetCellFalse != "" {
				shapes, err := createImprovedArrow(originCell, originCell, targetCellFalse, "rightConn", float64(shapeWidth), float64(shapeHeight), cellWidth, cellHeight, cellPadding)
				if err != nil {
					fmt.Printf("Error creating FALSE branch arrow: %v\n", err)
				} else {
					for _, shape := range shapes {
						file.AddShape("Sheet1", shape)
					}
				}
			}

			// TRUE branch arrow
			targetCellTrue := findCellByOrder(branchInfo.TrueTargetOrder, shapeOrders, shapeLocations)
			if targetCellTrue != "" {
				shapes, err := createImprovedArrow(originCell, originCell, targetCellTrue, "downConn", float64(shapeWidth), float64(shapeHeight), cellWidth, cellHeight, cellPadding)
				if err != nil {
					fmt.Printf("Error creating TRUE branch arrow: %v\n", err)
				} else {
					for _, shape := range shapes {
						file.AddShape("Sheet1", shape)
					}
				}
			}

		} else if i < len(shapeTypes)-1 {
			// Regular shape connections
			if _, isDecision := decisionBranches[i]; !isDecision {
				targetCell := shapeLocations[i+1]
				originCol, _, _ := excelize.CellNameToCoordinates(originCell)
				targetCol, _, _ := excelize.CellNameToCoordinates(targetCell)

				var orientation string
				if originCol == targetCol {
					orientation = "downConn"
				} else if targetCol > originCol {
					orientation = "rightConn"
				} else {
					orientation = "leftConn"
				}

				shapes, err := createImprovedArrow(originCell, originCell, targetCell, orientation, float64(shapeWidth), float64(shapeHeight), cellWidth, cellHeight, cellPadding)
				if err != nil {
					fmt.Printf("Error creating connector arrow: %v\n", err)
				} else {
					for _, shape := range shapes {
						file.AddShape("Sheet1", shape)
					}
				}
			}
		}
	}

	filename := fmt.Sprintf("improved_flowchart_%d.xlsx", rand.Intn(10000))
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	w.Header().Set("Content-Disposition", "attachment; filename="+filename)
	if err := file.Write(w); err != nil {
		http.Error(w, "Failed to generate file", http.StatusInternalServerError)
	}
}

// Helper functions
func pixelsToPoints(pixels float64) float64 {
	if pixels == 0 {
		return 0
	}
	return pixels * 3.0 / 4.0
}

func pixelsToCharUnits(pixels float64) float64 {
	if pixels <= 0 {
		return 0
	}
	return (pixels - 5) / 7
}

func getCellDifference(cell1, cell2 string) (width, height int, err error) {
	col1, row1, err := excelize.CellNameToCoordinates(cell1)
	if err != nil {
		return 0, 0, fmt.Errorf("invalid first cell name: %w", err)
	}

	col2, row2, err := excelize.CellNameToCoordinates(cell2)
	if err != nil {
		return 0, 0, fmt.Errorf("invalid second cell name: %w", err)
	}

	width = col2 - col1
	height = row2 - row1

	return width, height, nil
}
