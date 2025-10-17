// cmd/server/main.go
package main

import (
	"go_excelize/internal/app/handler"
	"go_excelize/internal/app/router"
	"go_excelize/internal/app/service"
	"log"
	"net/http"
)

func main() {
	// Services
	excelService := service.NewExcelService()

	// Handlers
	excelHandler := handler.NewExcelHandler(excelService)

	// Router
	r := router.NewRouter(excelHandler)

	// Start server
	log.Println("🚀 Server running at http://localhost:8080")
	if err := http.ListenAndServe(":8080", r); err != nil {
		log.Fatal(err)
	}
}
