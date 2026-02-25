package godocx_test

import (
	"os"
	"path/filepath"
	"testing"
	"time"

	godocx "github.com/falcomza/go-docx"
)

// TestNewBlank verifies that a blank document can be created from scratch.
func TestNewBlank(t *testing.T) {
	u, err := godocx.NewBlank()
	if err != nil {
		t.Fatalf("NewBlank() failed: %v", err)
	}
	defer u.Cleanup()

	// Verify tempDir exists
	if u.TempDir() == "" {
		t.Fatal("TempDir is empty")
	}

	// Verify core required files exist
	for _, relPath := range []string{
		"word/document.xml",
		"word/_rels/document.xml.rels",
		"[Content_Types].xml",
		"_rels/.rels",
		"docProps/core.xml",
		"docProps/app.xml",
	} {
		fullPath := filepath.Join(u.TempDir(), relPath)
		if _, err := os.Stat(fullPath); os.IsNotExist(err) {
			t.Errorf("missing required file: %s", relPath)
		}
	}
}

// TestNewBlankAddContent verifies that content can be inserted into a blank document.
func TestNewBlankAddContent(t *testing.T) {
	u, err := godocx.NewBlank()
	if err != nil {
		t.Fatalf("NewBlank() failed: %v", err)
	}
	defer u.Cleanup()

	err = u.InsertParagraph(godocx.ParagraphOptions{
		Text:     "Hello from a blank document!",
		Position: godocx.PositionEnd,
	})
	if err != nil {
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	text, err := u.GetText()
	if err != nil {
		t.Fatalf("GetText failed: %v", err)
	}
	if text == "" {
		t.Error("expected non-empty text after inserting paragraph")
	}
}

// TestNewBlankSaveAndReopen verifies that a blank document can be saved and reopened.
func TestNewBlankSaveAndReopen(t *testing.T) {
	outputDir := t.TempDir()
	outputPath := filepath.Join(outputDir, "blank_test.docx")

	u, err := godocx.NewBlank()
	if err != nil {
		t.Fatalf("NewBlank() failed: %v", err)
	}

	err = u.InsertParagraph(godocx.ParagraphOptions{
		Text:     "Saved from blank",
		Position: godocx.PositionEnd,
	})
	if err != nil {
		u.Cleanup()
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	if err := u.Save(outputPath); err != nil {
		u.Cleanup()
		t.Fatalf("Save failed: %v", err)
	}
	u.Cleanup()

	// Reopen the saved document
	u2, err := godocx.New(outputPath)
	if err != nil {
		t.Fatalf("Failed to reopen saved blank document: %v", err)
	}
	defer u2.Cleanup()

	text, err := u2.GetText()
	if err != nil {
		t.Fatalf("GetText on reopened doc failed: %v", err)
	}
	if text == "" {
		t.Error("expected text in reopened document")
	}
}

// TestNewFromBytes verifies creating an Updater from raw bytes.
func TestNewFromBytes(t *testing.T) {
	// First create a valid docx in memory
	u, err := godocx.NewBlank()
	if err != nil {
		t.Fatalf("NewBlank() failed: %v", err)
	}

	err = u.InsertParagraph(godocx.ParagraphOptions{
		Text:     "Bytes test content",
		Position: godocx.PositionEnd,
	})
	if err != nil {
		u.Cleanup()
		t.Fatalf("InsertParagraph failed: %v", err)
	}

	tmpPath := filepath.Join(t.TempDir(), "bytes_source.docx")
	if err := u.Save(tmpPath); err != nil {
		u.Cleanup()
		t.Fatalf("Save failed: %v", err)
	}
	u.Cleanup()

	// Read the file to bytes
	data, err := os.ReadFile(tmpPath)
	if err != nil {
		t.Fatalf("ReadFile failed: %v", err)
	}

	// Create from bytes
	u2, err := godocx.NewFromBytes(data)
	if err != nil {
		t.Fatalf("NewFromBytes() failed: %v", err)
	}
	defer u2.Cleanup()

	text, err := u2.GetText()
	if err != nil {
		t.Fatalf("GetText on bytes-loaded doc failed: %v", err)
	}
	if text == "" {
		t.Error("expected text in bytes-loaded document")
	}
}

// TestNewFromBytesEmpty verifies that NewFromBytes rejects empty data.
func TestNewFromBytesEmpty(t *testing.T) {
	_, err := godocx.NewFromBytes(nil)
	if err == nil {
		t.Error("expected error for nil data")
	}

	_, err = godocx.NewFromBytes([]byte{})
	if err == nil {
		t.Error("expected error for empty data")
	}
}

// TestGetAppProperties verifies reading app properties.
func TestGetAppProperties(t *testing.T) {
	u, err := godocx.NewBlank()
	if err != nil {
		t.Fatalf("NewBlank() failed: %v", err)
	}
	defer u.Cleanup()

	// Set app properties
	err = u.SetAppProperties(godocx.AppProperties{
		Company:     "TestCorp",
		Manager:     "Jane Doe",
		Application: "Microsoft Word",
		AppVersion:  "16.0000",
		Template:    "Report.dotm",
	})
	if err != nil {
		t.Fatalf("SetAppProperties failed: %v", err)
	}

	// Get app properties
	props, err := u.GetAppProperties()
	if err != nil {
		t.Fatalf("GetAppProperties failed: %v", err)
	}

	if props.Company != "TestCorp" {
		t.Errorf("Company: expected %q, got %q", "TestCorp", props.Company)
	}
	if props.Manager != "Jane Doe" {
		t.Errorf("Manager: expected %q, got %q", "Jane Doe", props.Manager)
	}
	if props.Template != "Report.dotm" {
		t.Errorf("Template: expected %q, got %q", "Report.dotm", props.Template)
	}
}

// TestGetAppPropertiesWithIntFields verifies integer fields in app properties.
func TestGetAppPropertiesWithIntFields(t *testing.T) {
	u, err := godocx.NewBlank()
	if err != nil {
		t.Fatalf("NewBlank() failed: %v", err)
	}
	defer u.Cleanup()

	err = u.SetAppProperties(godocx.AppProperties{
		Company:              "TestCo",
		Template:             "Normal.dotm",
		TotalTime:            45,
		Pages:                12,
		Words:                3500,
		Characters:           20000,
		CharactersWithSpaces: 23500,
		Lines:                180,
		Paragraphs:           42,
	})
	if err != nil {
		t.Fatalf("SetAppProperties failed: %v", err)
	}

	props, err := u.GetAppProperties()
	if err != nil {
		t.Fatalf("GetAppProperties failed: %v", err)
	}

	if props.TotalTime != 45 {
		t.Errorf("TotalTime: expected 45, got %d", props.TotalTime)
	}
	if props.Pages != 12 {
		t.Errorf("Pages: expected 12, got %d", props.Pages)
	}
	if props.Words != 3500 {
		t.Errorf("Words: expected 3500, got %d", props.Words)
	}
	if props.Characters != 20000 {
		t.Errorf("Characters: expected 20000, got %d", props.Characters)
	}
	if props.CharactersWithSpaces != 23500 {
		t.Errorf("CharactersWithSpaces: expected 23500, got %d", props.CharactersWithSpaces)
	}
	if props.Lines != 180 {
		t.Errorf("Lines: expected 180, got %d", props.Lines)
	}
	if props.Paragraphs != 42 {
		t.Errorf("Paragraphs: expected 42, got %d", props.Paragraphs)
	}
}

// TestGetCustomProperties verifies reading custom properties with various types.
func TestGetCustomProperties(t *testing.T) {
	u, err := godocx.NewBlank()
	if err != nil {
		t.Fatalf("NewBlank() failed: %v", err)
	}
	defer u.Cleanup()

	set := []godocx.CustomProperty{
		{Name: "ProjectName", Value: "Alpha"},
		{Name: "Version", Value: 42},
		{Name: "Budget", Value: 99.5},
		{Name: "Active", Value: true},
		{Name: "DueDate", Value: time.Date(2026, 6, 15, 0, 0, 0, 0, time.UTC)},
	}

	err = u.SetCustomProperties(set)
	if err != nil {
		t.Fatalf("SetCustomProperties failed: %v", err)
	}

	got, err := u.GetCustomProperties()
	if err != nil {
		t.Fatalf("GetCustomProperties failed: %v", err)
	}

	if len(got) != len(set) {
		t.Fatalf("expected %d properties, got %d", len(set), len(got))
	}

	// Verify string property
	if got[0].Name != "ProjectName" {
		t.Errorf("expected name %q, got %q", "ProjectName", got[0].Name)
	}
	if got[0].Value != "Alpha" {
		t.Errorf("expected value %q, got %v", "Alpha", got[0].Value)
	}

	// Verify int property
	if got[1].Name != "Version" {
		t.Errorf("expected name %q, got %q", "Version", got[1].Name)
	}
	if v, ok := got[1].Value.(int); !ok || v != 42 {
		t.Errorf("expected int value 42, got %v (%T)", got[1].Value, got[1].Value)
	}

	// Verify float property
	if v, ok := got[2].Value.(float64); !ok || v != 99.5 {
		t.Errorf("expected float value 99.5, got %v (%T)", got[2].Value, got[2].Value)
	}

	// Verify bool property
	if v, ok := got[3].Value.(bool); !ok || v != true {
		t.Errorf("expected bool value true, got %v (%T)", got[3].Value, got[3].Value)
	}

	// Verify date property
	if v, ok := got[4].Value.(time.Time); !ok {
		t.Errorf("expected time.Time, got %T", got[4].Value)
	} else if v.Year() != 2026 || v.Month() != 6 || v.Day() != 15 {
		t.Errorf("expected 2026-06-15, got %v", v)
	}
}

// TestGetCustomPropertiesEmpty verifies that GetCustomProperties returns nil
// when no custom.xml exists.
func TestGetCustomPropertiesEmpty(t *testing.T) {
	u, err := godocx.NewBlank()
	if err != nil {
		t.Fatalf("NewBlank() failed: %v", err)
	}
	defer u.Cleanup()

	got, err := u.GetCustomProperties()
	if err != nil {
		t.Fatalf("GetCustomProperties on blank doc failed: %v", err)
	}
	if got != nil {
		t.Errorf("expected nil for blank doc, got %v", got)
	}
}

// TestCorePropertiesContentStatus verifies the ContentStatus field.
func TestCorePropertiesContentStatus(t *testing.T) {
	u, err := godocx.NewBlank()
	if err != nil {
		t.Fatalf("NewBlank() failed: %v", err)
	}
	defer u.Cleanup()

	err = u.SetCoreProperties(godocx.CoreProperties{
		Title:         "Draft Document",
		ContentStatus: "Draft",
	})
	if err != nil {
		t.Fatalf("SetCoreProperties failed: %v", err)
	}

	props, err := u.GetCoreProperties()
	if err != nil {
		t.Fatalf("GetCoreProperties failed: %v", err)
	}

	if props.ContentStatus != "Draft" {
		t.Errorf("ContentStatus: expected %q, got %q", "Draft", props.ContentStatus)
	}
	if props.Title != "Draft Document" {
		t.Errorf("Title: expected %q, got %q", "Draft Document", props.Title)
	}
}

// TestNewBlankWithAllProperties verifies setting all property types on a blank document.
func TestNewBlankWithAllProperties(t *testing.T) {
	outputDir := t.TempDir()
	outputPath := filepath.Join(outputDir, "all_props.docx")

	u, err := godocx.NewBlank()
	if err != nil {
		t.Fatalf("NewBlank() failed: %v", err)
	}
	defer u.Cleanup()

	// Set core properties
	err = u.SetCoreProperties(godocx.CoreProperties{
		Title:          "Complete Props Test",
		Subject:        "Testing",
		Creator:        "Test Author",
		Keywords:       "test, complete",
		Description:    "All properties set",
		Category:       "Test",
		ContentStatus:  "Final",
		LastModifiedBy: "Reviewer",
		Revision:       "5",
		Created:        time.Date(2026, 1, 1, 0, 0, 0, 0, time.UTC),
		Modified:       time.Date(2026, 2, 1, 0, 0, 0, 0, time.UTC),
	})
	if err != nil {
		t.Fatalf("SetCoreProperties failed: %v", err)
	}

	// Set app properties
	err = u.SetAppProperties(godocx.AppProperties{
		Company:              "MegaCorp",
		Manager:              "Director",
		Application:          "go-docx",
		AppVersion:           "2.1.0",
		Template:             "Custom.dotm",
		HyperlinkBase:        "https://example.com",
		TotalTime:            120,
		Pages:                25,
		Words:                7500,
		Characters:           42000,
		CharactersWithSpaces: 49500,
		Lines:                350,
		Paragraphs:           80,
		DocSecurity:          0,
	})
	if err != nil {
		t.Fatalf("SetAppProperties failed: %v", err)
	}

	// Set custom properties
	err = u.SetCustomProperties([]godocx.CustomProperty{
		{Name: "Project", Value: "Phoenix"},
		{Name: "Sprint", Value: 14},
	})
	if err != nil {
		t.Fatalf("SetCustomProperties failed: %v", err)
	}

	// Save and reopen
	if err := u.Save(outputPath); err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	u2, err := godocx.New(outputPath)
	if err != nil {
		t.Fatalf("Reopen failed: %v", err)
	}
	defer u2.Cleanup()

	// Verify core
	coreProps, err := u2.GetCoreProperties()
	if err != nil {
		t.Fatalf("GetCoreProperties failed: %v", err)
	}
	if coreProps.Title != "Complete Props Test" {
		t.Errorf("Title: got %q", coreProps.Title)
	}
	if coreProps.ContentStatus != "Final" {
		t.Errorf("ContentStatus: got %q", coreProps.ContentStatus)
	}

	// Verify app
	appProps, err := u2.GetAppProperties()
	if err != nil {
		t.Fatalf("GetAppProperties failed: %v", err)
	}
	if appProps.Company != "MegaCorp" {
		t.Errorf("Company: got %q", appProps.Company)
	}
	if appProps.Template != "Custom.dotm" {
		t.Errorf("Template: got %q", appProps.Template)
	}
	if appProps.Pages != 25 {
		t.Errorf("Pages: got %d", appProps.Pages)
	}

	// Verify custom
	customProps, err := u2.GetCustomProperties()
	if err != nil {
		t.Fatalf("GetCustomProperties failed: %v", err)
	}
	if len(customProps) != 2 {
		t.Errorf("expected 2 custom props, got %d", len(customProps))
	}
}

// TestAppPropertiesHyperlinkBase verifies the HyperlinkBase field.
func TestAppPropertiesHyperlinkBase(t *testing.T) {
	u, err := godocx.NewBlank()
	if err != nil {
		t.Fatalf("NewBlank() failed: %v", err)
	}
	defer u.Cleanup()

	err = u.SetAppProperties(godocx.AppProperties{
		HyperlinkBase: "https://docs.example.com/reports",
	})
	if err != nil {
		t.Fatalf("SetAppProperties failed: %v", err)
	}

	props, err := u.GetAppProperties()
	if err != nil {
		t.Fatalf("GetAppProperties failed: %v", err)
	}

	if props.HyperlinkBase != "https://docs.example.com/reports" {
		t.Errorf("HyperlinkBase: expected %q, got %q", "https://docs.example.com/reports", props.HyperlinkBase)
	}
}

// TestAppPropertiesTemplate verifies the Template field for template assignment.
func TestAppPropertiesTemplate(t *testing.T) {
	u, err := godocx.NewBlank()
	if err != nil {
		t.Fatalf("NewBlank() failed: %v", err)
	}
	defer u.Cleanup()

	err = u.SetAppProperties(godocx.AppProperties{
		Template: "Corporate_Report.dotm",
	})
	if err != nil {
		t.Fatalf("SetAppProperties failed: %v", err)
	}

	props, err := u.GetAppProperties()
	if err != nil {
		t.Fatalf("GetAppProperties failed: %v", err)
	}

	if props.Template != "Corporate_Report.dotm" {
		t.Errorf("Template: expected %q, got %q", "Corporate_Report.dotm", props.Template)
	}
}
