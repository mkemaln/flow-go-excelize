// internal/router/router.go
package router

import (
	"go_excelize/internal/app/handler"
	"net/http"

	"github.com/go-chi/chi/v5"
	"github.com/go-chi/chi/v5/middleware"
)

func NewRouter(excelHandler *handler.ExcelHandler) http.Handler {
	r := chi.NewRouter()

	// Middlewares
	r.Use(middleware.Logger)
	r.Use(middleware.Recoverer)

	// Routes
	r.Get("/excel", excelHandler.GenerateExcel)

	return r
}
