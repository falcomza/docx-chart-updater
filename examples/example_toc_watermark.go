//go:build ignore

package main

import (
	"fmt"
	"log"
	"os"

	godocx "github.com/falcomza/go-docx"
)

// This example demonstrates advanced document features:
// - Custom styles
// - Table of Contents
// - Text watermark
// - Page numbering
// - Footnotes and endnotes
//
// Usage: go run examples/example_toc_watermark.go <template.docx> <output.docx>
func main() {
	if len(os.Args) < 3 {
		fmt.Println("Usage: go run examples/example_toc_watermark.go <template.docx> <output.docx>")
		os.Exit(1)
	}

	templateFile := os.Args[1]
	outputFile := os.Args[2]

	u, err := godocx.New(templateFile)
	if err != nil {
		log.Fatalf("Failed to open template: %v", err)
	}
	defer u.Cleanup()

	// ── 1. Add Custom Styles ──────────────────────────────────────

	err = u.AddStyles([]godocx.StyleDefinition{
		{
			ID:          "ReportTitle",
			Name:        "Report Title",
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
			ID:          "SectionHeading",
			Name:        "Section Heading",
			Type:        godocx.StyleTypeParagraph,
			BasedOn:     "Heading1",
			FontFamily:  "Calibri",
			FontSize:    32,
			Color:       "2E74B5",
			Bold:        true,
			KeepNext:    true,
			SpaceBefore: 360,
			SpaceAfter:  120,
		},
		{
			ID:         "SubHeading",
			Name:       "Sub Heading",
			Type:       godocx.StyleTypeParagraph,
			BasedOn:    "Heading2",
			FontFamily: "Calibri",
			FontSize:   26,
			Color:      "404040",
			Bold:       true,
			Italic:     true,
			KeepNext:   true,
			SpaceAfter: 80,
		},
		{
			ID:         "HighlightText",
			Name:       "Highlight Text",
			Type:       godocx.StyleTypeCharacter,
			Bold:       true,
			Color:      "C00000",
			FontFamily: "Calibri",
		},
	})
	if err != nil {
		log.Fatalf("Failed to add styles: %v", err)
	}
	fmt.Println("✓ Custom styles added")

	// ── 2. Add Title ─────────────────────────────────────────────

	err = u.InsertParagraph(godocx.ParagraphOptions{
		Text:     "Annual Performance Report 2026",
		Style:    "ReportTitle",
		Position: godocx.PositionBeginning,
	})
	if err != nil {
		log.Fatalf("Failed to add title: %v", err)
	}
	fmt.Println("✓ Title inserted")

	// ── 3. Insert Table of Contents ──────────────────────────────

	err = u.InsertTOC(godocx.TOCOptions{
		Title:         "Contents",
		OutlineLevels: "1-3",
		Position:      godocx.PositionAfterText,
		Anchor:        "Annual Performance Report 2026",
	})
	if err != nil {
		log.Fatalf("Failed to insert TOC: %v", err)
	}
	fmt.Println("✓ Table of Contents inserted")

	// ── 4. Add Document Content (Headings + Paragraphs) ──────────

	sections := []struct {
		heading string
		style   string
		body    string
	}{
		{
			heading: "Executive Summary",
			style:   "SectionHeading",
			body:    "This report provides a comprehensive overview of our annual performance metrics, strategic initiatives, and future outlook for the organization.",
		},
		{
			heading: "Key Findings",
			style:   "SectionHeading",
			body:    "Revenue increased by 23% year-over-year, driven by strong performance in the cloud services division. Customer satisfaction scores reached an all-time high of 94%.",
		},
		{
			heading: "Financial Overview",
			style:   "SubHeading",
			body:    "Total revenue reached $4.2 billion, with operating margins improving to 28%. Research and development investment grew to $890 million.",
		},
		{
			heading: "Strategic Initiatives",
			style:   "SectionHeading",
			body:    "Our digital transformation program delivered significant results across all business units. The AI-powered analytics platform was deployed to 15 enterprise customers.",
		},
		{
			heading: "Conclusion",
			style:   "SectionHeading",
			body:    "The organization demonstrated resilient growth and strong execution against strategic priorities. We remain well-positioned for continued success in the coming fiscal year.",
		},
	}

	for _, s := range sections {
		err = u.InsertParagraph(godocx.ParagraphOptions{
			Text:     s.heading,
			Style:    godocx.ParagraphStyle(s.style),
			Position: godocx.PositionEnd,
		})
		if err != nil {
			log.Fatalf("Failed to add heading %q: %v", s.heading, err)
		}

		err = u.InsertParagraph(godocx.ParagraphOptions{
			Text:     s.body,
			Position: godocx.PositionEnd,
		})
		if err != nil {
			log.Fatalf("Failed to add body for %q: %v", s.heading, err)
		}
	}
	fmt.Println("✓ Document content added")

	// ── 5. Add Footnotes ─────────────────────────────────────────

	err = u.InsertFootnote(godocx.FootnoteOptions{
		Text:   "Based on audited financial statements for fiscal year 2025-2026.",
		Anchor: "Revenue increased by 23%",
	})
	if err != nil {
		log.Fatalf("Failed to add footnote: %v", err)
	}

	err = u.InsertFootnote(godocx.FootnoteOptions{
		Text:   "Net Promoter Score methodology, surveyed Q4 2026.",
		Anchor: "satisfaction scores reached an all-time high",
	})
	if err != nil {
		log.Fatalf("Failed to add footnote: %v", err)
	}
	fmt.Println("✓ Footnotes inserted")

	// ── 6. Add Endnote ───────────────────────────────────────────

	err = u.InsertEndnote(godocx.EndnoteOptions{
		Text:   "For detailed methodology, see Appendix B of the full financial report.",
		Anchor: "Total revenue reached $4.2 billion",
	})
	if err != nil {
		log.Fatalf("Failed to add endnote: %v", err)
	}
	fmt.Println("✓ Endnote inserted")

	// ── 7. Set Page Numbering ────────────────────────────────────

	err = u.SetPageNumber(godocx.PageNumberOptions{
		Start:  1,
		Format: godocx.PageNumDecimal,
	})
	if err != nil {
		log.Fatalf("Failed to set page numbers: %v", err)
	}
	fmt.Println("✓ Page numbering configured")

	// ── 8. Add Watermark ─────────────────────────────────────────

	err = u.SetTextWatermark(godocx.WatermarkOptions{
		Text:       "DRAFT",
		FontFamily: "Calibri",
		Color:      "D0D0D0",
		Opacity:    0.35,
		Diagonal:   true,
	})
	if err != nil {
		log.Fatalf("Failed to add watermark: %v", err)
	}
	fmt.Println("✓ Watermark added")

	// ── 9. Add Footer with Page Numbers ──────────────────────────

	err = u.SetFooter(godocx.HeaderFooterContent{
		CenterText:       "Annual Performance Report 2026",
		PageNumber:       true,
		PageNumberFormat: "Page X of Y",
	}, godocx.DefaultFooterOptions())
	if err != nil {
		log.Fatalf("Failed to set footer: %v", err)
	}
	fmt.Println("✓ Footer with page numbers added")

	// ── 10. Mark TOC for Update ─────────────────────────────────

	err = u.UpdateTOC()
	if err != nil {
		log.Fatalf("Failed to update TOC: %v", err)
	}
	fmt.Println("✓ TOC marked for update")

	// ── Save ─────────────────────────────────────────────────────

	err = u.Save(outputFile)
	if err != nil {
		log.Fatalf("Failed to save: %v", err)
	}

	fmt.Printf("\n✅ Document saved to %s\n", outputFile)
	fmt.Println("   Open in Word and press Ctrl+A, F9 to update the Table of Contents.")
}
