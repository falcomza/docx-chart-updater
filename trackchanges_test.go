package godocx

import (
	"strings"
	"testing"
	"time"
)

func TestGenerateTrackedInsertXML(t *testing.T) {
	opts := TrackedInsertOptions{
		Text:   "Hello tracked world",
		Author: "Test Author",
		Date:   time.Date(2026, 1, 15, 10, 30, 0, 0, time.UTC),
		Style:  StyleNormal,
	}

	result := string(generateTrackedInsertXML(opts))

	if !strings.Contains(result, "<w:p>") {
		t.Error("expected <w:p> element")
	}
	if !strings.Contains(result, `<w:ins w:id="2" w:author="Test Author"`) {
		t.Error("expected w:ins element with author")
	}
	if !strings.Contains(result, "2026-01-15T10:30:00Z") {
		t.Error("expected date in ISO 8601 format")
	}
	if !strings.Contains(result, "Hello tracked world") {
		t.Error("expected text content")
	}
}

func TestGenerateTrackedInsertXML_WithFormatting(t *testing.T) {
	opts := TrackedInsertOptions{
		Text:   "Bold tracked text",
		Author: "Editor",
		Date:   time.Date(2026, 6, 1, 0, 0, 0, 0, time.UTC),
		Bold:   true,
		Italic: true,
	}

	result := string(generateTrackedInsertXML(opts))

	if !strings.Contains(result, "<w:b/>") {
		t.Error("expected bold formatting")
	}
	if !strings.Contains(result, "<w:i/>") {
		t.Error("expected italic formatting")
	}
	if !strings.Contains(result, "<w:ins") {
		t.Error("expected w:ins wrapper")
	}
}

func TestGenerateTrackedInsertXML_WithStyle(t *testing.T) {
	opts := TrackedInsertOptions{
		Text:   "Heading text",
		Author: "Author",
		Date:   time.Date(2026, 1, 1, 0, 0, 0, 0, time.UTC),
		Style:  StyleHeading1,
	}

	result := string(generateTrackedInsertXML(opts))

	if !strings.Contains(result, `<w:pStyle w:val="Heading1"/>`) {
		t.Error("expected Heading1 style")
	}
}

func TestConvertRunsToDeleted(t *testing.T) {
	para := `<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr>` +
		`<w:r><w:t>Hello world</w:t></w:r></w:p>`

	result := convertRunsToDeleted(para, "Reviewer", "2026-01-15T10:30:00Z")

	if !strings.Contains(result, "<w:del") {
		t.Error("expected w:del wrapper")
	}
	if !strings.Contains(result, "<w:delText") {
		t.Error("expected w:delText element")
	}
	if strings.Contains(result, "<w:t>") {
		t.Error("should not contain <w:t> after conversion")
	}
	if !strings.Contains(result, `w:author="Reviewer"`) {
		t.Error("expected author attribute")
	}
	if !strings.Contains(result, `w:date="2026-01-15T10:30:00Z"`) {
		t.Error("expected date attribute")
	}
}

func TestConvertRunsToDeleted_MultipleRuns(t *testing.T) {
	para := `<w:p><w:r><w:t>First</w:t></w:r>` +
		`<w:r><w:t>Second</w:t></w:r></w:p>`

	result := convertRunsToDeleted(para, "Author", "2026-01-01T00:00:00Z")

	// Should have two w:del wrappers
	count := strings.Count(result, "<w:del ")
	if count != 2 {
		t.Errorf("expected 2 w:del wrappers, got %d", count)
	}
}

func TestConvertRunsToDeleted_PreservesNonTextRuns(t *testing.T) {
	para := `<w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>` +
		`<w:footnoteReference w:id="1"/></w:r>` +
		`<w:r><w:t>Some text</w:t></w:r></w:p>`

	result := convertRunsToDeleted(para, "Author", "2026-01-01T00:00:00Z")

	// The footnote reference run should be preserved as-is
	if !strings.Contains(result, "<w:footnoteReference") {
		t.Error("expected footnote reference to be preserved")
	}
	// Only the text run should be wrapped in w:del
	count := strings.Count(result, "<w:del ")
	if count != 1 {
		t.Errorf("expected 1 w:del wrapper, got %d", count)
	}
}

func TestConvertRunsToDeleted_PreserveAttribute(t *testing.T) {
	para := `<w:p><w:r><w:t xml:space="preserve"> text </w:t></w:r></w:p>`

	result := convertRunsToDeleted(para, "Author", "2026-01-01T00:00:00Z")

	if !strings.Contains(result, `<w:delText xml:space="preserve">`) {
		t.Error("expected xml:space attribute to be preserved on delText")
	}
}

func TestInsertTrackedText_Validation(t *testing.T) {
	var u *Updater
	err := u.InsertTrackedText(TrackedInsertOptions{Text: "test"})
	if err == nil {
		t.Error("expected error for nil updater")
	}
}

func TestInsertTrackedText_EmptyText(t *testing.T) {
	u := &Updater{}
	err := u.InsertTrackedText(TrackedInsertOptions{})
	if err == nil || !strings.Contains(err.Error(), "text cannot be empty") {
		t.Error("expected 'text cannot be empty' error")
	}
}

func TestDeleteTrackedText_Validation(t *testing.T) {
	var u *Updater
	err := u.DeleteTrackedText(TrackedDeleteOptions{Anchor: "test"})
	if err == nil {
		t.Error("expected error for nil updater")
	}
}

func TestDeleteTrackedText_EmptyAnchor(t *testing.T) {
	u := &Updater{}
	err := u.DeleteTrackedText(TrackedDeleteOptions{})
	if err == nil || !strings.Contains(err.Error(), "anchor text cannot be empty") {
		t.Error("expected 'anchor text cannot be empty' error")
	}
}
