package godocx

import (
	"strings"
	"testing"
)

func TestInsertBreakAfterAnchor(t *testing.T) {
	doc := []byte(`<w:body><w:p><w:r><w:t>Chapter 1</w:t></w:r></w:p><w:p><w:r><w:t>Content</w:t></w:r></w:p></w:body>`)
	breakXML := []byte(`<w:p><w:r><w:br w:type="page"/></w:r></w:p>`)

	t.Run("insert after anchor", func(t *testing.T) {
		result, err := insertBreakAfterAnchor(doc, breakXML, "Chapter 1")
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		rs := string(result)
		// Break should appear after the first paragraph
		chapterIdx := strings.Index(rs, "</w:p>")
		breakIdx := strings.Index(rs, `w:type="page"`)
		if breakIdx < chapterIdx {
			t.Error("break should appear after the anchor paragraph")
		}
	})

	t.Run("anchor not found", func(t *testing.T) {
		_, err := insertBreakAfterAnchor(doc, breakXML, "nonexistent")
		if err == nil {
			t.Error("expected error for missing anchor")
		}
	})
}

func TestInsertBreakBeforeAnchor(t *testing.T) {
	doc := []byte(`<w:body><w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Chapter 2</w:t></w:r></w:p></w:body>`)
	breakXML := []byte(`<w:p><w:r><w:br w:type="page"/></w:r></w:p>`)

	t.Run("insert before anchor", func(t *testing.T) {
		result, err := insertBreakBeforeAnchor(doc, breakXML, "Chapter 2")
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		rs := string(result)
		// Break should appear before the second paragraph
		breakIdx := strings.Index(rs, `w:type="page"`)
		chapterIdx := strings.Index(rs, "Chapter 2")
		if breakIdx > chapterIdx {
			t.Error("break should appear before the anchor paragraph")
		}
	})

	t.Run("anchor not found", func(t *testing.T) {
		_, err := insertBreakBeforeAnchor(doc, breakXML, "nonexistent")
		if err == nil {
			t.Error("expected error for missing anchor")
		}
	})

	t.Run("insert before anchor with attributes", func(t *testing.T) {
		docAttrs := []byte(`<w:body><w:p rsidR="123"><w:r><w:t>Target</w:t></w:r></w:p></w:body>`)
		result, err := insertBreakBeforeAnchor(docAttrs, breakXML, "Target")
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		rs := string(result)
		if !strings.Contains(rs, `w:type="page"`) {
			t.Error("expected break in result")
		}
	})
}

func TestValidateSectionBreakType(t *testing.T) {
	validTypes := []SectionBreakType{
		SectionBreakNextPage,
		SectionBreakContinuous,
		SectionBreakEvenPage,
		SectionBreakOddPage,
	}
	for _, bt := range validTypes {
		if err := validateSectionBreakType(bt); err != nil {
			t.Errorf("expected %s to be valid, got error: %v", bt, err)
		}
	}

	if err := validateSectionBreakType("invalid"); err == nil {
		t.Error("expected error for invalid section break type")
	}
}

func TestInsertBreakAtPosition_AllPositions(t *testing.T) {
	doc := []byte(`<w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p><w:sectPr/></w:body>`)
	breakXML := generatePageBreakXML()

	t.Run("beginning", func(t *testing.T) {
		result, err := insertBreakAtPosition(doc, breakXML, BreakOptions{Position: PositionBeginning})
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if !strings.Contains(string(result), `w:type="page"`) {
			t.Error("expected break in result")
		}
	})

	t.Run("end", func(t *testing.T) {
		result, err := insertBreakAtPosition(doc, breakXML, BreakOptions{Position: PositionEnd})
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		if !strings.Contains(string(result), `w:type="page"`) {
			t.Error("expected break in result")
		}
	})

	t.Run("after text missing anchor", func(t *testing.T) {
		_, err := insertBreakAtPosition(doc, breakXML, BreakOptions{Position: PositionAfterText})
		if err == nil {
			t.Error("expected error for missing anchor")
		}
	})

	t.Run("before text missing anchor", func(t *testing.T) {
		_, err := insertBreakAtPosition(doc, breakXML, BreakOptions{Position: PositionBeforeText})
		if err == nil {
			t.Error("expected error for missing anchor")
		}
	})

	t.Run("invalid position", func(t *testing.T) {
		_, err := insertBreakAtPosition(doc, breakXML, BreakOptions{Position: InsertPosition(99)})
		if err == nil {
			t.Error("expected error for invalid position")
		}
	})
}

func TestGeneratePageBreakXML(t *testing.T) {
	result := string(generatePageBreakXML())
	if !strings.Contains(result, `w:type="page"`) {
		t.Error("expected page break type")
	}
	if !strings.Contains(result, "<w:p>") {
		t.Error("expected paragraph wrapper")
	}
}

func TestGenerateSectionBreakXML(t *testing.T) {
	t.Run("with page layout", func(t *testing.T) {
		layout := PageLayoutA4Landscape()
		result := string(generateSectionBreakXML(SectionBreakNextPage, layout))
		if !strings.Contains(result, `w:val="nextPage"`) {
			t.Error("expected nextPage section type")
		}
		if !strings.Contains(result, `w:orient="landscape"`) {
			t.Error("expected landscape orientation")
		}
	})

	t.Run("nil page layout uses defaults", func(t *testing.T) {
		result := string(generateSectionBreakXML(SectionBreakContinuous, nil))
		if !strings.Contains(result, `w:val="continuous"`) {
			t.Error("expected continuous section type")
		}
		if !strings.Contains(result, "<w:pgSz") {
			t.Error("expected default page size")
		}
	})
}
