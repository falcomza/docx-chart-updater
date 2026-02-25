//go:build ignore

package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"time"

	updater "github.com/falcomza/go-docx"
)

func main() {
	fmt.Println("Blank Document & Properties CRUD Example")
	fmt.Println("===========================================")

	// 1. Create a blank document from scratch (no template needed)
	fmt.Println("\n1. Creating blank document with NewBlank()...")
	u, err := updater.NewBlank()
	if err != nil {
		log.Fatal(err)
	}
	defer u.Cleanup()
	fmt.Println("   ✓ Blank document created")

	// 2. Set Core Properties (including new ContentStatus field)
	fmt.Println("\n2. Setting Core Properties...")
	coreProps := updater.CoreProperties{
		Title:          "Project Proposal 2026",
		Subject:        "New Product Launch Strategy",
		Creator:        "Engineering Team",
		Keywords:       "proposal, product, launch, strategy, 2026",
		Description:    "Detailed proposal for new product launch in Q3 2026",
		Category:       "Business Proposals",
		ContentStatus:  "Draft",
		LastModifiedBy: "Review Board",
		Revision:       "2",
		Created:        time.Date(2026, 3, 1, 9, 0, 0, 0, time.UTC),
		Modified:       time.Now(),
	}
	if err := u.SetCoreProperties(coreProps); err != nil {
		log.Fatal(err)
	}
	fmt.Println("   ✓ Title:", coreProps.Title)
	fmt.Println("   ✓ Status:", coreProps.ContentStatus)
	fmt.Println("   ✓ Author:", coreProps.Creator)

	// 3. Set Expanded App Properties (template, statistics, etc.)
	fmt.Println("\n3. Setting Application Properties...")
	appProps := updater.AppProperties{
		Company:              "Innovatech Solutions",
		Manager:              "Sarah Chen",
		Application:          "go-docx",
		AppVersion:           "2.1.0",
		Template:             "Corporate_Proposal.dotm",
		HyperlinkBase:        "https://docs.innovatech.com",
		TotalTime:            90,
		Pages:                15,
		Words:                4200,
		Characters:           24000,
		CharactersWithSpaces: 28200,
		Lines:                210,
		Paragraphs:           55,
	}
	if err := u.SetAppProperties(appProps); err != nil {
		log.Fatal(err)
	}
	fmt.Println("   ✓ Company:", appProps.Company)
	fmt.Println("   ✓ Template:", appProps.Template)
	fmt.Printf("   ✓ Statistics: %d pages, %d words, %d min editing\n",
		appProps.Pages, appProps.Words, appProps.TotalTime)

	// 4. Set Custom Properties with various types
	fmt.Println("\n4. Setting Custom Properties...")
	customProps := []updater.CustomProperty{
		{Name: "Department", Value: "Product Engineering"},
		{Name: "ProjectCode", Value: "PRJ-2026-LAUNCH"},
		{Name: "Priority", Value: "High"},
		{Name: "BudgetUSD", Value: 250000.75},
		{Name: "TeamSize", Value: 12},
		{Name: "IsApproved", Value: false},
		{Name: "Deadline", Value: time.Date(2026, 9, 30, 0, 0, 0, 0, time.UTC)},
	}
	if err := u.SetCustomProperties(customProps); err != nil {
		log.Fatal(err)
	}
	fmt.Printf("   ✓ Set %d custom properties\n", len(customProps))

	// 5. Read back Core Properties
	fmt.Println("\n5. Reading Core Properties...")
	readCore, err := u.GetCoreProperties()
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println("   Title:", readCore.Title)
	fmt.Println("   Author:", readCore.Creator)
	fmt.Println("   Status:", readCore.ContentStatus)
	fmt.Println("   Category:", readCore.Category)

	// 6. Read back App Properties
	fmt.Println("\n6. Reading Application Properties...")
	readApp, err := u.GetAppProperties()
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println("   Company:", readApp.Company)
	fmt.Println("   Template:", readApp.Template)
	fmt.Println("   HyperlinkBase:", readApp.HyperlinkBase)
	fmt.Printf("   Pages: %d, Words: %d, Lines: %d\n",
		readApp.Pages, readApp.Words, readApp.Lines)

	// 7. Read back Custom Properties
	fmt.Println("\n7. Reading Custom Properties...")
	readCustom, err := u.GetCustomProperties()
	if err != nil {
		log.Fatal(err)
	}
	for _, prop := range readCustom {
		fmt.Printf("   %s = %v (%T)\n", prop.Name, prop.Value, prop.Value)
	}

	// 8. Add content to the document
	fmt.Println("\n8. Adding content...")
	err = u.InsertParagraph(updater.ParagraphOptions{
		Text:     "Project Proposal: New Product Launch",
		Style:    updater.StyleTitle,
		Position: updater.PositionEnd,
	})
	if err != nil {
		log.Fatal(err)
	}

	err = u.InsertParagraph(updater.ParagraphOptions{
		Text:     "This document was created from scratch using NewBlank() with full property metadata.",
		Position: updater.PositionEnd,
	})
	if err != nil {
		log.Fatal(err)
	}

	err = u.InsertTable(updater.TableOptions{
		Columns: []updater.ColumnDefinition{
			{Title: "Property"},
			{Title: "Value"},
		},
		Rows: [][]string{
			{"Template", appProps.Template},
			{"Company", appProps.Company},
			{"Status", coreProps.ContentStatus},
			{"Author", coreProps.Creator},
		},
		Position:         updater.PositionEnd,
		TableStyle:       updater.TableStyleGridAccent1,
		HeaderBold:       true,
		HeaderBackground: "4472C4",
	})
	if err != nil {
		log.Fatal(err)
	}

	// 9. Save the document
	outputPath := "outputs/example_blank_with_properties.docx"
	fmt.Println("\n9. Saving document...")
	if err := u.Save(outputPath); err != nil {
		log.Fatal(err)
	}
	fmt.Println("   ✓ Saved:", outputPath)

	// 10. Demonstrate NewFromBytes (loading from raw bytes)
	fmt.Println("\n10. Demonstrating NewFromBytes()...")
	data, err := os.ReadFile(outputPath)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Printf("   Read %d bytes from saved document\n", len(data))

	u2, err := updater.NewFromBytes(data)
	if err != nil {
		log.Fatal(err)
	}
	defer u2.Cleanup()

	// Verify properties survived the round-trip
	verifyCore, err := u2.GetCoreProperties()
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println("   ✓ Round-trip Title:", verifyCore.Title)
	fmt.Println("   ✓ Round-trip Status:", verifyCore.ContentStatus)

	verifyApp, err := u2.GetAppProperties()
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println("   ✓ Round-trip Template:", verifyApp.Template)
	fmt.Println("   ✓ Round-trip Company:", verifyApp.Company)

	// Add more content and save under a new name
	err = u2.InsertParagraph(updater.ParagraphOptions{
		Text:     "This paragraph was added after loading from bytes.",
		Position: updater.PositionEnd,
		Bold:     true,
	})
	if err != nil {
		log.Fatal(err)
	}

	// Update status to Final
	err = u2.SetCoreProperties(updater.CoreProperties{
		Title:         verifyCore.Title,
		Creator:       verifyCore.Creator,
		ContentStatus: "Final",
	})
	if err != nil {
		log.Fatal(err)
	}

	finalPath := filepath.Join("outputs", "example_blank_final.docx")
	if err := u2.Save(finalPath); err != nil {
		log.Fatal(err)
	}
	fmt.Println("   ✓ Final saved:", finalPath)

	fmt.Println("\n✓ Complete! Documents saved to outputs/ directory.")
	fmt.Println("\nTo inspect properties in Word:")
	fmt.Println("  File > Info > Properties > Advanced Properties")
	fmt.Println("  • Summary tab: Title, Subject, Author, Keywords, Status")
	fmt.Println("  • Statistics tab: Pages, Words, Characters, Lines")
	fmt.Println("  • Custom tab: All custom name/value pairs")
}
