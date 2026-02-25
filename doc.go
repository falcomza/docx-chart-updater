// Package godocx provides a programmatic API for creating and modifying
// Microsoft Word (DOCX) documents without requiring Microsoft Office.
//
// # Quick Start
//
//	u, err := godocx.New("template.docx")
//	if err != nil {
//	    log.Fatal(err)
//	}
//	defer u.Cleanup()
//
//	u.UpdateChart(1, godocx.ChartData{
//	    Categories: []string{"Q1", "Q2", "Q3", "Q4"},
//	    Series: []godocx.SeriesData{
//	        {Name: "Revenue", Values: []float64{100, 150, 120, 180}},
//	    },
//	})
//
//	if err := u.Save("output.docx"); err != nil {
//	    log.Fatal(err)
//	}
//
// # Architecture
//
// go-docx extracts the DOCX ZIP archive to a temporary directory, manipulates
// the underlying OpenXML files directly, and repackages on [Updater.Save].
// Call [Updater.Cleanup] (typically via defer) to remove the temporary
// directory when done.
//
// # Creating Documents
//
// There are four ways to create or open a document:
//   - [New] — opens an existing DOCX file from disk
//   - [NewBlank] — creates a new empty document from scratch (no template needed)
//   - [NewFromBytes] — creates a document from raw bytes (e.g., uploaded template data)
//   - [NewFromReader] — creates a document from an [io.Reader]
//
// # Inserting Content
//
// All Insert* and Add* methods accept an [InsertPosition]:
//   - [PositionEnd] — appends to the document body
//   - [PositionBeginning] — prepends to the document body
//   - [PositionAfterText] — inserts after the paragraph containing Anchor text
//   - [PositionBeforeText] — inserts before the paragraph containing Anchor text
//
// # Document Properties
//
// Properties correspond to the Info panel and Advanced Properties dialog in Microsoft Word.
//
//   - [Updater.SetCoreProperties] / [Updater.GetCoreProperties] — title, author, keywords, status, dates
//   - [Updater.SetAppProperties] / [Updater.GetAppProperties] — company, manager, template, statistics
//   - [Updater.SetCustomProperties] / [Updater.GetCustomProperties] — user-defined name/value pairs
//
// The [AppProperties.Template] field assigns a template name (e.g., "Normal.dotm").
//
// # Chart Workflow
//
// Use [Updater.UpdateChart] to replace data in an existing chart template.
// Use [Updater.InsertChart] to create a new chart from scratch.
// Use [Updater.GetChartData] to read current chart categories and series.
package godocx
