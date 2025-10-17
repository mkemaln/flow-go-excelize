package service

import (
	"go_excelize/internal/app/model"
)

type ExcelService struct {
	excels []model.Input
}

// NewUserService initializes with mock data
func NewExcelService() *ExcelService {
	return &ExcelService{
		excels: []model.Input{
			{ID: 1, Name: "Alice"},
			{ID: 2, Name: "Bob"},
		},
	}
}
