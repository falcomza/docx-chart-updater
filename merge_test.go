package godocx

import (
	"strings"
	"testing"
)

func TestInjectTcPrElement_NoExistingTcPr(t *testing.T) {
	cell := `<w:tc><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:tc>`
	result := injectTcPrElement(cell, `<w:gridSpan w:val="3"/>`)
	if !strings.Contains(result, `<w:tcPr><w:gridSpan w:val="3"/></w:tcPr>`) {
		t.Errorf("expected tcPr block to be created, got: %s", result)
	}
}

func TestInjectTcPrElement_ExistingTcPr(t *testing.T) {
	cell := `<w:tc><w:tcPr><w:vAlign w:val="center"/></w:tcPr><w:p><w:r><w:t>Hi</w:t></w:r></w:p></w:tc>`
	result := injectTcPrElement(cell, `<w:vMerge w:val="restart"/>`)
	if !strings.Contains(result, `<w:tcPr><w:vMerge w:val="restart"/><w:vAlign`) {
		t.Errorf("expected vMerge injected into existing tcPr, got: %s", result)
	}
}

func TestInjectTcPrElement_SelfClosingTcPr(t *testing.T) {
	cell := `<w:tc><w:tcPr/><w:p><w:r><w:t>Hi</w:t></w:r></w:p></w:tc>`
	result := injectTcPrElement(cell, `<w:gridSpan w:val="2"/>`)
	if !strings.Contains(result, `<w:tcPr><w:gridSpan w:val="2"/></w:tcPr>`) {
		t.Errorf("expected self-closing tcPr to be expanded, got: %s", result)
	}
	if strings.Contains(result, "<w:tcPr/>") {
		t.Errorf("self-closing tcPr should have been replaced, got: %s", result)
	}
}

func TestMergeTableCellsHorizontal(t *testing.T) {
	docXML := `<?xml version="1.0"?><w:document><w:body>` +
		`<w:tbl><w:tr>` +
		`<w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>` +
		`<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>` +
		`<w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>` +
		`</w:tr></w:tbl>` +
		`</w:body></w:document>`

	result, err := mergeTableCellsHorizontal([]byte(docXML), 1, 1, 1, 2)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	s := string(result)
	if !strings.Contains(s, `<w:gridSpan w:val="2"/>`) {
		t.Errorf("expected gridSpan in result, got: %s", s)
	}
	// Should have only 2 cells (first merged + third)
	count := strings.Count(s, "<w:tc>") + strings.Count(s, "<w:tc ")
	if count != 2 {
		t.Errorf("expected 2 cells after merge, got %d", count)
	}
}

func TestMergeTableCellsVertical(t *testing.T) {
	docXML := `<?xml version="1.0"?><w:document><w:body>` +
		`<w:tbl>` +
		`<w:tr><w:tc><w:p><w:r><w:t>R1</w:t></w:r></w:p></w:tc></w:tr>` +
		`<w:tr><w:tc><w:p><w:r><w:t>R2</w:t></w:r></w:p></w:tc></w:tr>` +
		`<w:tr><w:tc><w:p><w:r><w:t>R3</w:t></w:r></w:p></w:tc></w:tr>` +
		`</w:tbl>` +
		`</w:body></w:document>`

	result, err := mergeTableCellsVertical([]byte(docXML), 1, 1, 3, 1)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	s := string(result)
	if !strings.Contains(s, `<w:vMerge w:val="restart"/>`) {
		t.Errorf("expected vMerge restart in first row, got: %s", s)
	}
	// Rows 2 and 3 should have continuation vMerge (no val attribute)
	if strings.Count(s, "<w:vMerge/>") != 2 {
		t.Errorf("expected 2 continuation vMerge elements, got: %s", s)
	}
}

func TestMergeHorizontal_InvalidParams(t *testing.T) {
	docXML := []byte(`<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`)

	_, err := mergeTableCellsHorizontal(docXML, 1, 1, 2, 1)
	if err == nil {
		t.Error("expected error for endCol <= startCol")
	}
}

func TestMergeVertical_InvalidParams(t *testing.T) {
	docXML := []byte(`<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`)

	_, err := mergeTableCellsVertical(docXML, 1, 2, 1, 1)
	if err == nil {
		t.Error("expected error for endRow <= startRow")
	}
}


