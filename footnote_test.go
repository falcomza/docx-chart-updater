package godocx

import (
	"bytes"
	"strings"
	"testing"
)

func TestGenerateFootnoteEntry(t *testing.T) {
	result := generateFootnoteEntry(1, "This is a footnote.")
	xml := string(result)

	if !strings.Contains(xml, `w:id="1"`) {
		t.Error("expected footnote id=1")
	}
	if !strings.Contains(xml, `<w:pStyle w:val="FootnoteText"/>`) {
		t.Error("expected FootnoteText style")
	}
	if !strings.Contains(xml, `<w:rStyle w:val="FootnoteReference"/>`) {
		t.Error("expected FootnoteReference run style")
	}
	if !strings.Contains(xml, "<w:footnoteRef/>") {
		t.Error("expected footnoteRef element")
	}
	if !strings.Contains(xml, "This is a footnote.") {
		t.Error("expected footnote text")
	}
}

func TestGenerateEndnoteEntry(t *testing.T) {
	result := generateEndnoteEntry(3, "This is an endnote.")
	xml := string(result)

	if !strings.Contains(xml, `w:id="3"`) {
		t.Error("expected endnote id=3")
	}
	if !strings.Contains(xml, `<w:pStyle w:val="EndnoteText"/>`) {
		t.Error("expected EndnoteText style")
	}
	if !strings.Contains(xml, `<w:rStyle w:val="EndnoteReference"/>`) {
		t.Error("expected EndnoteReference run style")
	}
	if !strings.Contains(xml, "<w:endnoteRef/>") {
		t.Error("expected endnoteRef element")
	}
	if !strings.Contains(xml, "This is an endnote.") {
		t.Error("expected endnote text")
	}
}

func TestGenerateInitialFootnotesXML(t *testing.T) {
	result := generateInitialFootnotesXML()
	xml := string(result)

	if !strings.Contains(xml, `<w:footnotes`) {
		t.Error("expected <w:footnotes> root element")
	}
	if !strings.Contains(xml, `w:type="separator"`) {
		t.Error("expected separator footnote")
	}
	if !strings.Contains(xml, `w:type="continuationSeparator"`) {
		t.Error("expected continuation separator footnote")
	}
	if !strings.Contains(xml, `w:id="-1"`) {
		t.Error("expected separator with id=-1")
	}
	if !strings.Contains(xml, `w:id="0"`) {
		t.Error("expected continuation separator with id=0")
	}
	if !strings.Contains(xml, "</w:footnotes>") {
		t.Error("expected closing tag")
	}
}

func TestGenerateInitialEndnotesXML(t *testing.T) {
	result := generateInitialEndnotesXML()
	xml := string(result)

	if !strings.Contains(xml, `<w:endnotes`) {
		t.Error("expected <w:endnotes> root element")
	}
	if !strings.Contains(xml, `w:type="separator"`) {
		t.Error("expected separator endnote")
	}
	if !strings.Contains(xml, `w:type="continuationSeparator"`) {
		t.Error("expected continuation separator endnote")
	}
	if !strings.Contains(xml, "</w:endnotes>") {
		t.Error("expected closing tag")
	}
}

func TestGetNextNoteID_Empty(t *testing.T) {
	// Only has separator footnotes (id=-1, id=0)
	raw := generateInitialFootnotesXML()

	nextID := getNextNoteID(raw, "footnote")

	if nextID != 1 {
		t.Errorf("expected next ID 1, got %d", nextID)
	}
}

func TestGetNextNoteID_WithExisting(t *testing.T) {
	// Create footnotes.xml with existing footnotes
	var buf bytes.Buffer
	buf.Write(generateInitialFootnotesXML())

	// Manually inject a footnote with id=1 and id=2
	raw := bytes.Replace(buf.Bytes(),
		[]byte("</w:footnotes>"),
		[]byte(`<w:footnote w:id="1"><w:p/></w:footnote>`+
			`<w:footnote w:id="2"><w:p/></w:footnote>`+
			`</w:footnotes>`),
		1)

	nextID := getNextNoteID(raw, "footnote")

	if nextID != 3 {
		t.Errorf("expected next ID 3, got %d", nextID)
	}
}

func TestFootnoteEntryXMLEscaping(t *testing.T) {
	result := generateFootnoteEntry(1, `Text with "quotes" & <special> chars`)
	xml := string(result)

	if !strings.Contains(xml, "&amp;") {
		t.Error("expected escaped ampersand")
	}
	if !strings.Contains(xml, "&lt;special&gt;") {
		t.Error("expected escaped angle brackets")
	}
	if !strings.Contains(xml, "&quot;quotes&quot;") {
		t.Error("expected escaped quotes")
	}
}
