package godocx

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"
)

// TrackedInsertOptions defines options for inserting text with revision tracking.
type TrackedInsertOptions struct {
	// Text to insert (required)
	Text string

	// Author of the revision (default: "Author")
	Author string

	// Date of the revision (default: current time)
	Date time.Time

	// Position where to insert the tracked text
	Position InsertPosition

	// Anchor text for position-based insertion
	Anchor string

	// Style for the inserted paragraph (default: Normal)
	Style ParagraphStyle

	// Text formatting
	Bold      bool
	Italic    bool
	Underline bool
}

// TrackedDeleteOptions defines options for marking existing text as deleted.
type TrackedDeleteOptions struct {
	// Anchor is the text in the document to mark as deleted (required).
	// The entire paragraph containing this text will be marked.
	Anchor string

	// Author of the deletion revision (default: "Author")
	Author string

	// Date of the deletion revision (default: current time)
	Date time.Time
}

// InsertTrackedText inserts a new paragraph with revision tracking.
// The inserted text appears as a tracked insertion (green underline in Word)
// that can be accepted or rejected by the reviewer.
func (u *Updater) InsertTrackedText(opts TrackedInsertOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}
	if opts.Text == "" {
		return fmt.Errorf("text cannot be empty")
	}
	if opts.Author == "" {
		opts.Author = "Author"
	}
	if opts.Date.IsZero() {
		opts.Date = time.Now()
	}
	if opts.Style == "" {
		opts.Style = StyleNormal
	}

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	startID := getNextRevisionID(raw)
	trackedXML := generateTrackedInsertXMLWithID(opts, startID)

	updated, err := insertTrackedAtPosition(raw, trackedXML, opts)
	if err != nil {
		return fmt.Errorf("insert tracked text: %w", err)
	}

	if err := os.WriteFile(docPath, updated, 0o644); err != nil {
		return fmt.Errorf("write document.xml: %w", err)
	}

	return nil
}

// DeleteTrackedText marks the paragraph containing the anchor text as a
// tracked deletion. The text appears as struck-through red text in Word
// and can be accepted or rejected by the reviewer.
func (u *Updater) DeleteTrackedText(opts TrackedDeleteOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}
	if opts.Anchor == "" {
		return fmt.Errorf("anchor text cannot be empty")
	}
	if opts.Author == "" {
		opts.Author = "Author"
	}
	if opts.Date.IsZero() {
		opts.Date = time.Now()
	}

	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return fmt.Errorf("read document.xml: %w", err)
	}

	startID := getNextRevisionID(raw)
	updated, err := markParagraphAsDeleted(raw, opts, startID)
	if err != nil {
		return fmt.Errorf("mark tracked deletion: %w", err)
	}

	if err := os.WriteFile(docPath, updated, 0o644); err != nil {
		return fmt.Errorf("write document.xml: %w", err)
	}

	return nil
}

// revisionIDPattern matches w:id attributes on revision elements (w:ins, w:del, etc.)
var revisionIDPattern = regexp.MustCompile(`w:id="(\d+)"`)

// getNextRevisionID scans the document XML for the highest existing revision ID
// and returns the next available one.
func getNextRevisionID(docXML []byte) int {
	matches := revisionIDPattern.FindAllSubmatch(docXML, -1)
	maxID := 0
	for _, m := range matches {
		if len(m) > 1 {
			if id, err := strconv.Atoi(string(m[1])); err == nil && id > maxID {
				maxID = id
			}
		}
	}
	return maxID + 1
}

// generateTrackedInsertXML creates a paragraph wrapped in w:ins revision markup
// using hardcoded IDs 1 and 2 (for unit tests only; use generateTrackedInsertXMLWithID for documents).
func generateTrackedInsertXML(opts TrackedInsertOptions) []byte {
	return generateTrackedInsertXMLWithID(opts, 1)
}

// generateTrackedInsertXMLWithID creates a paragraph wrapped in w:ins revision markup
// using the given starting revision ID.
func generateTrackedInsertXMLWithID(opts TrackedInsertOptions, startID int) []byte {
	var buf bytes.Buffer

	dateStr := opts.Date.UTC().Format(time.RFC3339)

	buf.WriteString("<w:p>")

	// Paragraph properties
	buf.WriteString("<w:pPr>")
	if opts.Style != StyleNormal {
		buf.WriteString(fmt.Sprintf(`<w:pStyle w:val="%s"/>`, xmlEscape(string(opts.Style))))
	}
	// Mark the paragraph properties themselves as an insertion
	buf.WriteString(fmt.Sprintf(
		`<w:rPr><w:ins w:id="%d" w:author="%s" w:date="%s"/></w:rPr>`,
		startID, xmlEscape(opts.Author), dateStr))
	buf.WriteString("</w:pPr>")

	// The text run wrapped in w:ins
	buf.WriteString(fmt.Sprintf(
		`<w:ins w:id="%d" w:author="%s" w:date="%s">`,
		startID+1, xmlEscape(opts.Author), dateStr))

	buf.WriteString("<w:r>")

	// Run properties
	hasFormatting := opts.Bold || opts.Italic || opts.Underline
	if hasFormatting {
		buf.WriteString("<w:rPr>")
		if opts.Bold {
			buf.WriteString("<w:b/>")
		}
		if opts.Italic {
			buf.WriteString("<w:i/>")
		}
		if opts.Underline {
			buf.WriteString(`<w:u w:val="single"/>`)
		}
		buf.WriteString("</w:rPr>")
	}

	writeRunTextWithControls(&buf, opts.Text)

	buf.WriteString("</w:r>")
	buf.WriteString("</w:ins>")
	buf.WriteString("</w:p>")

	return buf.Bytes()
}

// insertTrackedAtPosition inserts the tracked XML at the specified position.
func insertTrackedAtPosition(docXML, trackedXML []byte, opts TrackedInsertOptions) ([]byte, error) {
	// Reuse the same position logic as paragraphs
	pOpts := ParagraphOptions{
		Position: opts.Position,
		Anchor:   opts.Anchor,
	}
	return insertParagraphAtPosition(docXML, trackedXML, pOpts)
}

// markParagraphAsDeleted wraps the text runs of the paragraph containing
// the anchor text in w:del markup.
func markParagraphAsDeleted(docXML []byte, opts TrackedDeleteOptions, startID int) ([]byte, error) {
	paraStart, paraEnd, err := findParagraphRangeByAnchor(docXML, opts.Anchor)
	if err != nil {
		return nil, fmt.Errorf("find anchor: %w", err)
	}

	paraXML := string(docXML[paraStart:paraEnd])
	dateStr := opts.Date.UTC().Format(time.RFC3339)

	// We need to convert all <w:r>...<w:t>text</w:t>...</w:r> runs into
	// <w:del><w:r>...<w:delText>text</w:delText>...</w:r></w:del>
	modified := convertRunsToDeletedWithID(paraXML, opts.Author, dateStr, startID)

	result := make([]byte, 0, len(docXML)-len(paraXML)+len(modified))
	result = append(result, docXML[:paraStart]...)
	result = append(result, []byte(modified)...)
	result = append(result, docXML[paraEnd:]...)

	return result, nil
}

// convertRunsToDeleted converts text runs in a paragraph to deleted runs
// using sequential IDs starting from 1 (for unit tests).
func convertRunsToDeleted(paraXML, author, dateStr string) string {
	return convertRunsToDeletedWithID(paraXML, author, dateStr, 1)
}

// convertRunsToDeletedWithID converts text runs in a paragraph to deleted runs.
// It wraps each <w:r> containing <w:t> in <w:del> and replaces <w:t> with <w:delText>.
func convertRunsToDeletedWithID(paraXML, author, dateStr string, startID int) string {
	var result strings.Builder
	delID := startID

	pos := 0
	for pos < len(paraXML) {
		// Find next <w:r> that's not inside <w:rPr>
		runStart := strings.Index(paraXML[pos:], "<w:r>")
		runStartAttr := strings.Index(paraXML[pos:], "<w:r ")

		// Find the earliest run start
		nextRun := -1
		if runStart >= 0 {
			nextRun = runStart
		}
		if runStartAttr >= 0 && (nextRun < 0 || runStartAttr < nextRun) {
			nextRun = runStartAttr
		}

		if nextRun < 0 {
			result.WriteString(paraXML[pos:])
			break
		}
		nextRun += pos

		// Write content before this run
		result.WriteString(paraXML[pos:nextRun])

		// Find the end of this run
		runEnd := strings.Index(paraXML[nextRun:], "</w:r>")
		if runEnd < 0 {
			result.WriteString(paraXML[nextRun:])
			break
		}
		runEnd += nextRun + len("</w:r>")

		runContent := paraXML[nextRun:runEnd]

		// Check if this run contains text (<w:t>)
		if strings.Contains(runContent, "<w:t") {
			// Convert <w:t> to <w:delText> and wrap in <w:del>
			delRun := runContent
			delRun = strings.ReplaceAll(delRun, "<w:t>", `<w:delText xml:space="preserve">`)
			delRun = strings.ReplaceAll(delRun, "<w:t ", "<w:delText ")
			delRun = strings.ReplaceAll(delRun, "</w:t>", "</w:delText>")

			result.WriteString(fmt.Sprintf(
				`<w:del w:id="%d" w:author="%s" w:date="%s">`,
				delID, xmlEscape(author), dateStr))
			result.WriteString(delRun)
			result.WriteString("</w:del>")
			delID++
		} else {
			// Keep non-text runs as-is (e.g., footnote references)
			result.WriteString(runContent)
		}

		pos = runEnd
	}

	return result.String()
}
