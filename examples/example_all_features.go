//go:build ignore

package main

import (
	"fmt"
	"log"
	"os"
	"time"

	godocx "github.com/falcomza/go-docx"
)

// This example demonstrates every major feature of the go-docx library:
//   - Document properties (core, app, custom)
//   - Custom styles
//   - Table of Contents
//   - Headings, paragraphs, text formatting
//   - Bullet and numbered lists
//   - Tables with cell merging
//   - Charts (column, line, scatter)
//   - Bookmarks and hyperlinks
//   - Footnotes and endnotes
//   - Comments
//   - Track changes (insertions and deletions)
//   - Text watermark
//   - Page numbering
//   - Headers and footers
//   - Page breaks and section breaks
//   - Page layout
//
// Usage:
//
//	go run examples/example_all_features.go <template.docx> <output.docx>
func main() {
	if len(os.Args) < 3 {
		fmt.Println("Usage: go run examples/example_all_features.go <template.docx> <output.docx>")
		os.Exit(1)
	}

	templateFile := os.Args[1]
	outputFile := os.Args[2]

	u, err := godocx.New(templateFile)
	if err != nil {
		log.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// ── 1. Document Properties ──────────────────────────────────

	err = u.SetCoreProperties(godocx.CoreProperties{
		Title:       "Complete Feature Showcase",
		Subject:     "go-docx Library Demonstration",
		Creator:     "go-docx Example",
		Keywords:    "docx, golang, automation, report",
		Description: "Demonstrates every feature of the go-docx library",
		Category:    "Demo",
	})
	checkErr(err, "set core properties")

	err = u.SetAppProperties(godocx.AppProperties{
		Company:     "ACME Corporation",
		Application: "go-docx v2.1.0",
	})
	checkErr(err, "set app properties")

	err = u.SetCustomProperties([]godocx.CustomProperty{
		{Name: "Department", Value: "Engineering"},
		{Name: "Version", Value: "2.1.0"},
		{Name: "Reviewed", Value: "true"},
	})
	checkErr(err, "set custom properties")
	fmt.Println("✓ Document properties set")

	// ── 2. Custom Styles ────────────────────────────────────────

	err = u.AddStyles([]godocx.StyleDefinition{
		{
			ID:          "DocTitle",
			Name:        "Document Title",
			Type:        godocx.StyleTypeParagraph,
			BasedOn:     "Normal",
			FontFamily:  "Calibri",
			FontSize:    56,
			Color:       "1F4E79",
			Bold:        true,
			Alignment:   godocx.ParagraphAlignCenter,
			SpaceAfter:  240,
			SpaceBefore: 480,
		},
		{
			ID:          "ChapterHeading",
			Name:        "Chapter Heading",
			Type:        godocx.StyleTypeParagraph,
			BasedOn:     "Heading1",
			FontFamily:  "Calibri",
			FontSize:    36,
			Color:       "2E74B5",
			Bold:        true,
			KeepNext:    true,
			SpaceBefore: 360,
			SpaceAfter:  120,
		},
		{
			ID:         "Emphasis",
			Name:       "Strong Emphasis",
			Type:       godocx.StyleTypeCharacter,
			Bold:       true,
			Italic:     true,
			Color:      "C00000",
			FontFamily: "Calibri",
		},
	})
	checkErr(err, "add styles")
	fmt.Println("✓ Custom styles added")

	// ── 3. Title ────────────────────────────────────────────────

	err = u.InsertParagraph(godocx.ParagraphOptions{
		Text:     "go-docx Complete Feature Showcase",
		Style:    "DocTitle",
		Position: godocx.PositionBeginning,
	})
	checkErr(err, "add title")
	fmt.Println("✓ Title inserted")

	// ── 4. Table of Contents ────────────────────────────────────

	err = u.InsertTOC(godocx.TOCOptions{
		Title:         "Table of Contents",
		OutlineLevels: "1-3",
		Position:      godocx.PositionAfterText,
		Anchor:        "go-docx Complete Feature Showcase",
	})
	checkErr(err, "insert TOC")
	fmt.Println("✓ Table of Contents inserted")

	// ── 5. Introduction Section ─────────────────────────────────

	err = u.AddHeading(1, "Introduction", godocx.PositionEnd)
	checkErr(err, "add introduction heading")

	err = u.AddText(
		"This document demonstrates every feature available in the go-docx library. "+
			"Each section showcases a different capability, from basic text formatting "+
			"to advanced features like track changes and chart generation.",
		godocx.PositionEnd)
	checkErr(err, "add intro text")

	// ── 6. Create a Bookmark ────────────────────────────────────

	err = u.CreateBookmarkWithText("data_section", "Data Analysis", godocx.BookmarkOptions{
		Position: godocx.PositionEnd,
		Style:    godocx.StyleHeading1,
	})
	checkErr(err, "create bookmark")
	fmt.Println("✓ Bookmark created")

	// ── 7. Tables with Cell Merging ─────────────────────────────

	err = u.AddHeading(2, "Quarterly Performance", godocx.PositionEnd)
	checkErr(err, "add table heading")

	err = u.InsertTable(godocx.TableOptions{
		Position: godocx.PositionEnd,
		Columns: []godocx.ColumnDefinition{
			{Title: "Region", Width: 2000, Bold: true},
			{Title: "Q1", Width: 1200},
			{Title: "Q2", Width: 1200},
			{Title: "Q3", Width: 1200},
			{Title: "Q4", Width: 1200},
			{Title: "Total", Width: 1500, Bold: true},
		},
		Rows: [][]string{
			{"North America", "$250K", "$275K", "$290K", "$310K", "$1,125K"},
			{"Europe", "$180K", "$195K", "$205K", "$220K", "$800K"},
			{"Asia Pacific", "$150K", "$165K", "$180K", "$200K", "$695K"},
			{"Total", "$580K", "$635K", "$675K", "$730K", "$2,620K"},
		},
		HeaderBackground:  "2E75B6",
		HeaderBold:        true,
		AlternateRowColor: "D9E2F3",
		TableStyle:        godocx.TableStyleProfessional,
		RepeatHeader:      true,
		Caption: &godocx.CaptionOptions{
			Type:        godocx.CaptionTable,
			Description: "Regional quarterly performance",
			AutoNumber:  true,
		},
	})
	checkErr(err, "insert table")
	fmt.Println("✓ Table inserted")

	// ── 8. Charts ───────────────────────────────────────────────

	err = u.AddHeading(2, "Visual Analytics", godocx.PositionEnd)
	checkErr(err, "add chart heading")

	// Column chart
	err = u.InsertChart(godocx.ChartOptions{
		Position:          godocx.PositionEnd,
		ChartKind:         godocx.ChartKindColumn,
		Title:             "Revenue by Quarter",
		CategoryAxisTitle: "Quarter",
		ValueAxisTitle:    "Revenue (USD)",
		Categories:        []string{"Q1", "Q2", "Q3", "Q4"},
		Series: []godocx.SeriesOptions{
			{Name: "2025", Values: []float64{520000, 580000, 610000, 670000}},
			{Name: "2026", Values: []float64{580000, 635000, 675000, 730000}},
		},
		ShowLegend:     true,
		LegendPosition: "b",
		Caption: &godocx.CaptionOptions{
			Type:        godocx.CaptionFigure,
			Description: "Year-over-year revenue comparison",
			AutoNumber:  true,
		},
	})
	checkErr(err, "insert column chart")

	// Line chart
	err = u.InsertChart(godocx.ChartOptions{
		Position:          godocx.PositionEnd,
		ChartKind:         godocx.ChartKindLine,
		Title:             "Customer Growth Trend",
		CategoryAxisTitle: "Month",
		ValueAxisTitle:    "Customers",
		Categories:        []string{"Jan", "Feb", "Mar", "Apr", "May", "Jun"},
		Series: []godocx.SeriesOptions{
			{Name: "Enterprise", Values: []float64{120, 135, 142, 158, 170, 185}},
			{Name: "SMB", Values: []float64{450, 480, 520, 540, 575, 610}},
		},
		ShowLegend: true,
	})
	checkErr(err, "insert line chart")
	fmt.Println("✓ Charts inserted")

	// ── 9. Lists ────────────────────────────────────────────────

	err = u.AddHeading(1, "Strategic Priorities", godocx.PositionEnd)
	checkErr(err, "add list heading")

	err = u.AddText("Our key priorities for the coming year:", godocx.PositionEnd)
	checkErr(err, "add list intro")

	err = u.AddBulletList([]string{
		"Expand cloud services platform",
		"Increase market share in Asia Pacific",
		"Launch AI-powered analytics product",
		"Improve customer retention to 95%",
	}, 0, godocx.PositionEnd)
	checkErr(err, "add bullet list")

	err = u.AddText("Implementation timeline:", godocx.PositionEnd)
	checkErr(err, "add numbered list intro")

	err = u.AddNumberedList([]string{
		"Phase 1: Research and planning (Q1)",
		"Phase 2: Development and testing (Q2-Q3)",
		"Phase 3: Pilot deployment (Q3)",
		"Phase 4: Full rollout (Q4)",
	}, 0, godocx.PositionEnd)
	checkErr(err, "add numbered list")
	fmt.Println("✓ Lists added")

	// ── 10. Page Break ──────────────────────────────────────────

	err = u.InsertPageBreak(godocx.BreakOptions{
		Position: godocx.PositionEnd,
	})
	checkErr(err, "insert page break")

	// ── 11. Hyperlinks ─────────────────────────────────────────

	err = u.AddHeading(1, "Resources", godocx.PositionEnd)
	checkErr(err, "add resources heading")

	err = u.AddText("For more information, visit our website:", godocx.PositionEnd)
	checkErr(err, "add hyperlink intro")

	err = u.InsertHyperlink("Visit go-docx on GitHub", "https://github.com/falcomza/go-docx", godocx.HyperlinkOptions{
		Position:  godocx.PositionEnd,
		Color:     "0563C1",
		Underline: true,
		Tooltip:   "Open go-docx repository",
	})
	checkErr(err, "insert hyperlink")

	// Internal link to bookmark
	err = u.InsertInternalLink("Jump to Data Analysis section", "data_section", godocx.HyperlinkOptions{
		Position:  godocx.PositionEnd,
		Color:     "0563C1",
		Underline: true,
	})
	checkErr(err, "insert internal link")
	fmt.Println("✓ Hyperlinks added")

	// ── 12. Track Changes ──────────────────────────────────────

	err = u.AddHeading(1, "Review Section", godocx.PositionEnd)
	checkErr(err, "add review heading")

	err = u.AddText("This paragraph will be marked for deletion by a reviewer.", godocx.PositionEnd)
	checkErr(err, "add text for deletion")

	// Insert tracked text (appears as green underline in Word)
	err = u.InsertTrackedText(godocx.TrackedInsertOptions{
		Text:     "This paragraph was added during review and needs approval.",
		Author:   "Jane Reviewer",
		Date:     time.Now(),
		Position: godocx.PositionEnd,
	})
	checkErr(err, "insert tracked text")

	// Mark existing text as deleted (appears as red strikethrough in Word)
	err = u.DeleteTrackedText(godocx.TrackedDeleteOptions{
		Anchor: "marked for deletion by a reviewer",
		Author: "Jane Reviewer",
		Date:   time.Now(),
	})
	checkErr(err, "delete tracked text")
	fmt.Println("✓ Track changes applied")

	// ── 13. Comments ────────────────────────────────────────────

	err = u.InsertComment(godocx.CommentOptions{
		Text:     "Great progress on cloud expansion!",
		Author:   "Manager",
		Initials: "M",
		Anchor:   "Expand cloud services platform",
	})
	checkErr(err, "insert comment")
	fmt.Println("✓ Comment added")

	// ── 14. Footnotes and Endnotes ──────────────────────────────

	err = u.InsertFootnote(godocx.FootnoteOptions{
		Text:   "Based on audited financial statements for FY 2025-2026.",
		Anchor: "Revenue by Quarter",
	})
	checkErr(err, "insert footnote")

	err = u.InsertEndnote(godocx.EndnoteOptions{
		Text:   "For detailed methodology, see the full financial report appendix.",
		Anchor: "Customer Growth Trend",
	})
	checkErr(err, "insert endnote")
	fmt.Println("✓ Footnotes and endnotes added")

	// ── 15. Page Numbering ─────────────────────────────────────

	err = u.SetPageNumber(godocx.PageNumberOptions{
		Start:  1,
		Format: godocx.PageNumDecimal,
	})
	checkErr(err, "set page numbers")
	fmt.Println("✓ Page numbering configured")

	// ── 16. Watermark ──────────────────────────────────────────

	err = u.SetTextWatermark(godocx.WatermarkOptions{
		Text:       "CONFIDENTIAL",
		FontFamily: "Calibri",
		Color:      "E0E0E0",
		Opacity:    0.3,
		Diagonal:   true,
	})
	checkErr(err, "set watermark")
	fmt.Println("✓ Watermark added")

	// ── 17. Header and Footer ──────────────────────────────────

	err = u.SetHeader(godocx.HeaderFooterContent{
		CenterText: "CONFIDENTIAL — go-docx Feature Showcase",
	}, godocx.DefaultHeaderOptions())
	checkErr(err, "set header")

	err = u.SetFooter(godocx.HeaderFooterContent{
		LeftText:         "© 2026 ACME Corp",
		CenterText:       "Complete Feature Showcase",
		PageNumber:       true,
		PageNumberFormat: "Page X of Y",
	}, godocx.DefaultFooterOptions())
	checkErr(err, "set footer")
	fmt.Println("✓ Header and footer set")

	// ── 18. Mark TOC for Update ────────────────────────────────

	err = u.UpdateTOC()
	checkErr(err, "update TOC")
	fmt.Println("✓ TOC marked for update")

	// ── Save ────────────────────────────────────────────────────

	err = u.Save(outputFile)
	checkErr(err, "save")

	fmt.Printf("\n✅ Document saved to %s\n", outputFile)
	fmt.Println("   Features demonstrated:")
	fmt.Println("   • Document properties (core, app, custom)")
	fmt.Println("   • Custom styles (paragraph + character)")
	fmt.Println("   • Table of Contents with auto-update")
	fmt.Println("   • Headings, paragraphs, text formatting")
	fmt.Println("   • Bullet and numbered lists")
	fmt.Println("   • Tables with styling and captions")
	fmt.Println("   • Column and line charts with captions")
	fmt.Println("   • Bookmarks and hyperlinks (external + internal)")
	fmt.Println("   • Footnotes and endnotes")
	fmt.Println("   • Comments")
	fmt.Println("   • Track changes (insertions + deletions)")
	fmt.Println("   • Text watermark")
	fmt.Println("   • Page numbering")
	fmt.Println("   • Headers and footers with page numbers")
	fmt.Println("   • Page breaks")
	fmt.Println("")
	fmt.Println("   Open in Word and press Ctrl+A, F9 to update the Table of Contents.")
}

func checkErr(err error, action string) {
	if err != nil {
		log.Fatalf("Failed to %s: %v", action, err)
	}
}
