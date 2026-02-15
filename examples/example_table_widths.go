package main

import (
	"fmt"
	"log"

	docxchartupdater "github.com/falcomza/docx-chart-updater/src"
)

func main() {
	// Open the template document
	updater, err := docxchartupdater.New("./templates/docx_template.docx")
	if err != nil {
		log.Fatalf("Failed to open template: %v", err)
	}
	defer updater.Cleanup()

	// Add title
	err = updater.AddHeading(1, "Table Width Configuration Examples", docxchartupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add title: %v", err)
	}

	// Example 1: Default (100% width - spans between margins)
	err = updater.AddHeading(2, "1. Default Width (100% - Between Margins)", docxchartupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add heading: %v", err)
	}

	err = updater.InsertTable(docxchartupdater.TableOptions{
		Position: docxchartupdater.PositionEnd,
		Columns: []docxchartupdater.ColumnDefinition{
			{Title: "Product"},
			{Title: "Description"},
			{Title: "Price"},
		},
		Rows: [][]string{
			{"Laptop", "High-performance business laptop", "$1,299"},
			{"Mouse", "Wireless ergonomic mouse", "$49"},
			{"Keyboard", "Mechanical RGB keyboard", "$129"},
		},
		HeaderBold:        true,
		HeaderBackground:  "4472C4",
		HeaderAlignment:   docxchartupdater.CellAlignCenter,
		AlternateRowColor: "E7E6E6",
		BorderStyle:       docxchartupdater.BorderSingle,
	})
	if err != nil {
		log.Fatalf("Failed to insert default width table: %v", err)
	}

	err = updater.AddText("↑ This table spans 100% of the width between left and right margins (default behavior).", docxchartupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add description: %v", err)
	}

	// Example 2: 50% width table
	err = updater.AddHeading(2, "2. Custom Width (50%)", docxchartupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add heading: %v", err)
	}

	err = updater.InsertTable(docxchartupdater.TableOptions{
		Position: docxchartupdater.PositionEnd,
		Columns: []docxchartupdater.ColumnDefinition{
			{Title: "Code"},
			{Title: "Status"},
		},
		Rows: [][]string{
			{"A001", "✓ Active"},
			{"B002", "✓ Active"},
			{"C003", "⊗ Inactive"},
		},
		TableWidthType:    docxchartupdater.TableWidthPercentage,
		TableWidth:        2500,                               // 50% (5000 = 100%)
		HeaderBold:        true,
		HeaderBackground:  "70AD47",
		HeaderAlignment:   docxchartupdater.CellAlignCenter,
		AlternateRowColor: "E2EFD9",
		BorderStyle:       docxchartupdater.BorderSingle,
		TableAlignment:    docxchartupdater.AlignCenter,
	})
	if err != nil {
		log.Fatalf("Failed to insert 50%% width table: %v", err)
	}

	err = updater.AddText("↑ This table is 50% of the available width and centered.", docxchartupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add description: %v", err)
	}

	// Example 3: 75% width table
	err = updater.AddHeading(2, "3. Custom Width (75%)", docxchartupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add heading: %v", err)
	}

	err = updater.InsertTable(docxchartupdater.TableOptions{
		Position: docxchartupdater.PositionEnd,
		Columns: []docxchartupdater.ColumnDefinition{
			{Title: "Quarter", Alignment: docxchartupdater.CellAlignCenter},
			{Title: "Revenue", Alignment: docxchartupdater.CellAlignRight},
			{Title: "Growth", Alignment: docxchartupdater.CellAlignRight},
		},
		Rows: [][]string{
			{"Q1 2026", "$250,000", "+12%"},
			{"Q2 2026", "$280,000", "+15%"},
			{"Q3 2026", "$310,000", "+18%"},
		},
		TableWidthType:    docxchartupdater.TableWidthPercentage,
		TableWidth:        3750,                               // 75% (5000 = 100%)
		HeaderBold:        true,
		HeaderBackground:  "2E75B5",
		HeaderAlignment:   docxchartupdater.CellAlignCenter,
		AlternateRowColor: "DEEBF7",
		BorderStyle:       docxchartupdater.BorderSingle,
	})
	if err != nil {
		log.Fatalf("Failed to insert 75%% width table: %v", err)
	}

	err = updater.AddText("↑ This table is 75% of the available width.", docxchartupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add description: %v", err)
	}

	// Example 4: Fixed width in twips (5 inches = 7200 twips)
	err = updater.AddHeading(2, "4. Fixed Width (5 inches / 7200 twips)", docxchartupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add heading: %v", err)
	}

	err = updater.InsertTable(docxchartupdater.TableOptions{
		Position: docxchartupdater.PositionEnd,
		Columns: []docxchartupdater.ColumnDefinition{
			{Title: "Name"},
			{Title: "Email"},
		},
		Rows: [][]string{
			{"John Doe", "john@example.com"},
			{"Jane Smith", "jane@example.com"},
		},
		TableWidthType:    docxchartupdater.TableWidthFixed,
		TableWidth:        7200,                               // 5 inches
		HeaderBold:        true,
		HeaderBackground:  "C65911",
		HeaderAlignment:   docxchartupdater.CellAlignCenter,
		AlternateRowColor: "FCE4D6",
		BorderStyle:       docxchartupdater.BorderSingle,
	})
	if err != nil {
		log.Fatalf("Failed to insert fixed width table: %v", err)
	}

	err = updater.AddText("↑ This table has a fixed width of exactly 5 inches (7200 twips).", docxchartupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add description: %v", err)
	}

	// Example 5: Auto width (fits to content)
	err = updater.AddHeading(2, "5. Auto Width (Fits Content)", docxchartupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add heading: %v", err)
	}

	err = updater.InsertTable(docxchartupdater.TableOptions{
		Position: docxchartupdater.PositionEnd,
		Columns: []docxchartupdater.ColumnDefinition{
			{Title: "#"},
			{Title: "Item"},
		},
		Rows: [][]string{
			{"1", "Short"},
			{"2", "Text"},
			{"3", "Here"},
		},
		TableWidthType:    docxchartupdater.TableWidthAuto,
		HeaderBold:        true,
		HeaderBackground:  "7030A0",
		HeaderAlignment:   docxchartupdater.CellAlignCenter,
		AlternateRowColor: "E9D8F4",
		BorderStyle:       docxchartupdater.BorderSingle,
	})
	if err != nil {
		log.Fatalf("Failed to insert auto width table: %v", err)
	}

	err = updater.AddText("↑ This table auto-fits to its content width.", docxchartupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add description: %v", err)
	}

	// Add summary section
	err = updater.AddHeading(2, "Width Configuration Summary", docxchartupdater.PositionEnd)
	if err != nil {
		log.Fatalf("Failed to add summary heading: %v", err)
	}

	err = updater.InsertTable(docxchartupdater.TableOptions{
		Position: docxchartupdater.PositionEnd,
		Columns: []docxchartupdater.ColumnDefinition{
			{Title: "Type", Alignment: docxchartupdater.CellAlignLeft},
			{Title: "Value", Alignment: docxchartupdater.CellAlignLeft},
			{Title: "Description", Alignment: docxchartupdater.CellAlignLeft},
		},
		Rows: [][]string{
			{"Percentage (default)", "5000 = 100%", "Spans between margins (default)"},
			{"Percentage", "2500 = 50%", "Half of available width"},
			{"Percentage", "3750 = 75%", "Three quarters of width"},
			{"Fixed (twips)", "7200", "5 inches (1440 twips per inch)"},
			{"Auto", "n/a", "Fits to content"},
		},
		HeaderBold:        true,
		HeaderBackground:  "44546A",
		HeaderAlignment:   docxchartupdater.CellAlignCenter,
		AlternateRowColor: "D6DCE4",
		BorderStyle:       docxchartupdater.BorderSingle,
		RepeatHeader:      true,
		RowStyle: docxchartupdater.CellStyle{
			FontSize: 18, // 9pt
		},
	})
	if err != nil {
		log.Fatalf("Failed to insert summary table: %v", err)
	}

	// Save the document
	outputPath := "./outputs/table_width_examples.docx"
	if err := updater.Save(outputPath); err != nil {
		log.Fatalf("Failed to save document: %v", err)
	}

	fmt.Println("✅ SUCCESS!")
	fmt.Printf("📄 Output saved to: %s\n", outputPath)
	fmt.Println("\nCreated examples:")
	fmt.Println("  • Default (100% width - between margins)")
	fmt.Println("  • 50% width (centered)")
	fmt.Println("  • 75% width")
	fmt.Println("  • Fixed 5 inches (7200 twips)")
	fmt.Println("  • Auto width (fits content)")
	fmt.Println("  • Summary table with all options")
}
