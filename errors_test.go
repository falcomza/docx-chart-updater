package godocx

import (
	"errors"
	"fmt"
	"strings"
	"testing"
)

func TestDocxError_Error(t *testing.T) {
	tests := []struct {
		name     string
		err      *DocxError
		contains []string
	}{
		{
			name: "with wrapped error",
			err: &DocxError{
				Code:    ErrCodeChartNotFound,
				Message: "chart not found",
				Err:     fmt.Errorf("index out of range"),
			},
			contains: []string{"CHART_NOT_FOUND", "chart not found", "index out of range"},
		},
		{
			name: "without wrapped error",
			err: &DocxError{
				Code:    ErrCodeValidation,
				Message: "field required",
			},
			contains: []string{"VALIDATION", "field required"},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := tt.err.Error()
			for _, s := range tt.contains {
				if !strings.Contains(result, s) {
					t.Errorf("Error() = %q, expected to contain %q", result, s)
				}
			}
		})
	}
}

func TestDocxError_Unwrap(t *testing.T) {
	inner := fmt.Errorf("inner error")
	err := &DocxError{
		Code:    ErrCodeXMLParse,
		Message: "parse failed",
		Err:     inner,
	}

	if err.Unwrap() != inner {
		t.Error("Unwrap() did not return the inner error")
	}

	// Verify it works with errors.Is/errors.As
	if !errors.Is(err, inner) {
		t.Error("errors.Is should find the inner error")
	}
}

func TestDocxError_Unwrap_Nil(t *testing.T) {
	err := &DocxError{
		Code:    ErrCodeValidation,
		Message: "no inner",
	}

	if err.Unwrap() != nil {
		t.Error("Unwrap() should return nil when no inner error")
	}
}

func TestDocxError_WithContext(t *testing.T) {
	err := &DocxError{
		Code:    ErrCodeChartNotFound,
		Message: "chart not found",
	}

	result := err.WithContext("index", 5)
	if result != err {
		t.Error("WithContext should return the same error")
	}
	if err.Context == nil {
		t.Fatal("Context should be initialized")
	}
	if err.Context["index"] != 5 {
		t.Errorf("expected context index=5, got %v", err.Context["index"])
	}

	// Add more context
	err.WithContext("file", "chart1.xml")
	if err.Context["file"] != "chart1.xml" {
		t.Errorf("expected context file=chart1.xml, got %v", err.Context["file"])
	}
}

func TestDocxError_WithContext_PreExistingContext(t *testing.T) {
	err := &DocxError{
		Code:    ErrCodeValidation,
		Message: "test",
		Context: map[string]any{"existing": true},
	}

	err.WithContext("new", "value")
	if err.Context["existing"] != true {
		t.Error("existing context should be preserved")
	}
	if err.Context["new"] != "value" {
		t.Error("new context should be added")
	}
}

func TestNewChartNotFoundError(t *testing.T) {
	err := NewChartNotFoundError(3)
	if err == nil {
		t.Fatal("expected non-nil error")
	}
	if !strings.Contains(err.Error(), "CHART_NOT_FOUND") {
		t.Errorf("expected CHART_NOT_FOUND in error: %s", err.Error())
	}
	var docxErr *DocxError
	if !errors.As(err, &docxErr) {
		t.Fatal("expected DocxError type")
	}
	if docxErr.Context["index"] != 3 {
		t.Errorf("expected index=3, got %v", docxErr.Context["index"])
	}
}

func TestNewInvalidChartDataError(t *testing.T) {
	err := NewInvalidChartDataError("missing categories")
	if !strings.Contains(err.Error(), "missing categories") {
		t.Errorf("expected reason in error: %s", err.Error())
	}
}

func TestNewImageNotFoundError(t *testing.T) {
	err := NewImageNotFoundError("/path/to/img.png")
	var docxErr *DocxError
	if !errors.As(err, &docxErr) {
		t.Fatal("expected DocxError")
	}
	if docxErr.Code != ErrCodeImageNotFound {
		t.Errorf("expected code IMAGE_NOT_FOUND, got %s", docxErr.Code)
	}
	if docxErr.Context["path"] != "/path/to/img.png" {
		t.Error("expected path context")
	}
}

func TestNewImageFormatError(t *testing.T) {
	err := NewImageFormatError("svg")
	var docxErr *DocxError
	if !errors.As(err, &docxErr) {
		t.Fatal("expected DocxError")
	}
	if docxErr.Code != ErrCodeImageFormat {
		t.Errorf("expected code IMAGE_FORMAT, got %s", docxErr.Code)
	}
	if docxErr.Context["format"] != "svg" {
		t.Error("expected format context")
	}
}

func TestNewTextNotFoundError(t *testing.T) {
	err := NewTextNotFoundError("missing text")
	var docxErr *DocxError
	if !errors.As(err, &docxErr) {
		t.Fatal("expected DocxError")
	}
	if docxErr.Code != ErrCodeTextNotFound {
		t.Errorf("expected code TEXT_NOT_FOUND, got %s", docxErr.Code)
	}
}

func TestNewInvalidRegexError(t *testing.T) {
	inner := fmt.Errorf("bad pattern")
	err := NewInvalidRegexError("[invalid", inner)
	var docxErr *DocxError
	if !errors.As(err, &docxErr) {
		t.Fatal("expected DocxError")
	}
	if docxErr.Err != inner {
		t.Error("expected wrapped error")
	}
	if docxErr.Context["pattern"] != "[invalid" {
		t.Error("expected pattern context")
	}
}

func TestNewXMLParseError(t *testing.T) {
	inner := fmt.Errorf("xml broken")
	err := NewXMLParseError("chart1.xml", inner)
	var docxErr *DocxError
	if !errors.As(err, &docxErr) {
		t.Fatal("expected DocxError")
	}
	if docxErr.Code != ErrCodeXMLParse {
		t.Errorf("expected code XML_PARSE, got %s", docxErr.Code)
	}
}

func TestNewXMLWriteError(t *testing.T) {
	inner := fmt.Errorf("disk full")
	err := NewXMLWriteError("doc.xml", inner)
	var docxErr *DocxError
	if !errors.As(err, &docxErr) {
		t.Fatal("expected DocxError")
	}
	if docxErr.Code != ErrCodeXMLWrite {
		t.Errorf("expected code XML_WRITE, got %s", docxErr.Code)
	}
}

func TestNewRelationshipError(t *testing.T) {
	inner := fmt.Errorf("missing rel")
	err := NewRelationshipError("bad relationship", inner)
	var docxErr *DocxError
	if !errors.As(err, &docxErr) {
		t.Fatal("expected DocxError")
	}
	if docxErr.Code != ErrCodeRelationship {
		t.Errorf("expected code RELATIONSHIP, got %s", docxErr.Code)
	}
}

func TestNewValidationError(t *testing.T) {
	err := NewValidationError("title", "cannot be empty")
	var docxErr *DocxError
	if !errors.As(err, &docxErr) {
		t.Fatal("expected DocxError")
	}
	if docxErr.Context["field"] != "title" {
		t.Error("expected field context")
	}
}

func TestNewFileNotFoundError(t *testing.T) {
	err := NewFileNotFoundError("/missing/file.docx")
	var docxErr *DocxError
	if !errors.As(err, &docxErr) {
		t.Fatal("expected DocxError")
	}
	if docxErr.Code != ErrCodeFileNotFound {
		t.Errorf("expected FILE_NOT_FOUND, got %s", docxErr.Code)
	}
}

func TestNewInvalidFileError(t *testing.T) {
	inner := fmt.Errorf("not a zip")
	err := NewInvalidFileError("invalid docx", inner)
	var docxErr *DocxError
	if !errors.As(err, &docxErr) {
		t.Fatal("expected DocxError")
	}
	if docxErr.Code != ErrCodeInvalidFile {
		t.Errorf("expected INVALID_FILE, got %s", docxErr.Code)
	}
}

func TestNewHyperlinkError(t *testing.T) {
	inner := fmt.Errorf("creation failed")
	err := NewHyperlinkError("bad link", inner)
	var docxErr *DocxError
	if !errors.As(err, &docxErr) {
		t.Fatal("expected DocxError")
	}
	if docxErr.Code != ErrCodeHyperlinkCreation {
		t.Errorf("expected HYPERLINK_CREATION, got %s", docxErr.Code)
	}
}

func TestNewInvalidURLError(t *testing.T) {
	err := NewInvalidURLError("not-a-url")
	var docxErr *DocxError
	if !errors.As(err, &docxErr) {
		t.Fatal("expected DocxError")
	}
	if docxErr.Context["url"] != "not-a-url" {
		t.Error("expected url context")
	}
}

func TestNewHeaderFooterError(t *testing.T) {
	inner := fmt.Errorf("write failed")
	err := NewHeaderFooterError("header error", inner)
	var docxErr *DocxError
	if !errors.As(err, &docxErr) {
		t.Fatal("expected DocxError")
	}
	if docxErr.Code != ErrCodeHeaderFooter {
		t.Errorf("expected HEADER_FOOTER, got %s", docxErr.Code)
	}
}
