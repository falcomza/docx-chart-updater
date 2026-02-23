package godocx

import (
	"strings"
	"testing"
)

// --- styles.go coverage ---

func TestGenerateStyleXML_ParagraphFull(t *testing.T) {
	def := StyleDefinition{
		ID:            "CustomPara",
		Name:          "Custom Paragraph",
		Type:          StyleTypeParagraph,
		BasedOn:       "Normal",
		NextStyle:     "Normal",
		FontFamily:    "Calibri",
		FontSize:      28,
		Color:         "1F4E79",
		Bold:          true,
		Italic:        true,
		Underline:     true,
		Strikethrough: true,
		AllCaps:       true,
		SmallCaps:     true,
		Alignment:     ParagraphAlignCenter,
		SpaceBefore:   240,
		SpaceAfter:    120,
		LineSpacing:   480,
		IndentLeft:    720,
		IndentRight:   360,
		IndentFirst:   360,
		KeepNext:      true,
		KeepLines:     true,
		PageBreakBef:  true,
		OutlineLevel:  1,
	}

	result := string(generateStyleXML(def))

	assertContains(t, result, `w:styleId="CustomPara"`)
	assertContains(t, result, `w:type="paragraph"`)
	assertContains(t, result, `<w:name w:val="Custom Paragraph"/>`)
	assertContains(t, result, `<w:basedOn w:val="Normal"/>`)
	assertContains(t, result, `<w:next w:val="Normal"/>`)
	assertContains(t, result, "<w:b/>")
	assertContains(t, result, "<w:i/>")
	assertContains(t, result, `<w:u w:val="single"/>`)
	assertContains(t, result, "<w:strike/>")
	assertContains(t, result, "<w:caps/>")
	assertContains(t, result, "<w:smallCaps/>")
	assertContains(t, result, `<w:jc w:val="center"/>`)
	assertContains(t, result, `<w:sz w:val="28"/>`)
	assertContains(t, result, `<w:color w:val="1F4E79"/>`)
	assertContains(t, result, "<w:keepNext/>")
	assertContains(t, result, "<w:keepLines/>")
	assertContains(t, result, "<w:pageBreakBefore/>")
}

func TestGenerateStyleXML_CharacterMinimal(t *testing.T) {
	def := StyleDefinition{
		ID:   "Emph",
		Name: "Emphasis",
		Type: StyleTypeCharacter,
		Bold: true,
	}

	result := string(generateStyleXML(def))

	assertContains(t, result, `w:type="character"`)
	assertContains(t, result, "<w:b/>")
	// Character styles should NOT have pPr
	if strings.Contains(result, "<w:pPr>") {
		t.Error("character style should not have paragraph properties")
	}
}

func TestAddStyle_Validation(t *testing.T) {
	var u *Updater
	err := u.AddStyle(StyleDefinition{ID: "test"})
	if err == nil {
		t.Error("expected error for nil updater")
	}

	u = &Updater{}
	err = u.AddStyle(StyleDefinition{})
	if err == nil {
		t.Error("expected error for empty ID")
	}
}

// --- footnote.go coverage ---

func TestInsertFootnote_Validation(t *testing.T) {
	var u *Updater
	err := u.InsertFootnote(FootnoteOptions{Text: "test", Anchor: "test"})
	if err == nil {
		t.Error("expected error for nil updater")
	}
}

func TestInsertEndnote_Validation(t *testing.T) {
	var u *Updater
	err := u.InsertEndnote(EndnoteOptions{Text: "test", Anchor: "test"})
	if err == nil {
		t.Error("expected error for nil updater")
	}
}

// --- watermark.go coverage ---

func TestSetTextWatermark_Validation(t *testing.T) {
	var u *Updater
	err := u.SetTextWatermark(WatermarkOptions{Text: "test"})
	if err == nil {
		t.Error("expected error for nil updater")
	}
}

// --- pagenumber.go coverage ---

func TestSetPageNumberInSectPr(t *testing.T) {
	tests := []struct {
		name     string
		input    string
		opts     PageNumberOptions
		contains []string
	}{
		{
			name:  "decimal format",
			input: `<w:body><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body>`,
			opts:  PageNumberOptions{Start: 1, Format: PageNumDecimal},
			contains: []string{`w:start="1"`, `w:fmt="decimal"`},
		},
		{
			name:  "upper roman",
			input: `<w:body><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body>`,
			opts:  PageNumberOptions{Start: 5, Format: PageNumUpperRoman},
			contains: []string{`w:start="5"`, `w:fmt="upperRoman"`},
		},
		{
			name:  "lower roman",
			input: `<w:body><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body>`,
			opts:  PageNumberOptions{Start: 10, Format: PageNumLowerRoman},
			contains: []string{`w:start="10"`, `w:fmt="lowerRoman"`},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result, err := setPageNumberInSectPr([]byte(tt.input), tt.opts)
			if err != nil {
				t.Fatalf("unexpected error: %v", err)
			}
			rs := string(result)
			for _, s := range tt.contains {
				if !strings.Contains(rs, s) {
					t.Errorf("expected %q in result: %s", s, rs)
				}
			}
		})
	}
}

func TestSetPageNumber_Validation(t *testing.T) {
	var u *Updater
	err := u.SetPageNumber(PageNumberOptions{Start: 1})
	if err == nil {
		t.Error("expected error for nil updater")
	}
}

// --- trackchanges.go coverage ---

func TestGetNextRevisionID(t *testing.T) {
	tests := []struct {
		name string
		input string
		want  int
	}{
		{
			name:  "no revisions",
			input: `<w:body><w:p/></w:body>`,
			want:  1,
		},
		{
			name:  "existing revisions",
			input: `<w:ins w:id="5"/><w:del w:id="3"/>`,
			want:  6,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := getNextRevisionID([]byte(tt.input))
			if got != tt.want {
				t.Errorf("got %d, want %d", got, tt.want)
			}
		})
	}
}

// --- toc.go coverage ---

func TestInsertTOC_Validation(t *testing.T) {
	var u *Updater
	err := u.InsertTOC(TOCOptions{})
	if err == nil {
		t.Error("expected error for nil updater")
	}
}

func TestUpdateTOC_Validation(t *testing.T) {
	var u *Updater
	err := u.UpdateTOC()
	if err == nil {
		t.Error("expected error for nil updater")
	}
}

func TestGetTOCEntries_Validation(t *testing.T) {
	var u *Updater
	_, err := u.GetTOCEntries()
	if err == nil {
		t.Error("expected error for nil updater")
	}
}

// --- merge.go coverage ---

func TestMergeTableCellsHorizontal_Validation(t *testing.T) {
	var u *Updater
	err := u.MergeTableCellsHorizontal(1, 1, 1, 2)
	if err == nil {
		t.Error("expected error for nil updater")
	}

	u = &Updater{}
	err = u.MergeTableCellsHorizontal(0, 1, 1, 2)
	if err == nil {
		t.Error("expected error for tableIndex < 1")
	}
	err = u.MergeTableCellsHorizontal(1, 0, 1, 2)
	if err == nil {
		t.Error("expected error for row < 1")
	}
	err = u.MergeTableCellsHorizontal(1, 1, 0, 2)
	if err == nil {
		t.Error("expected error for startCol < 1")
	}
	err = u.MergeTableCellsHorizontal(1, 1, 2, 1)
	if err == nil {
		t.Error("expected error for endCol <= startCol")
	}
}

func TestMergeTableCellsVertical_Validation(t *testing.T) {
	var u *Updater
	err := u.MergeTableCellsVertical(1, 1, 2, 1)
	if err == nil {
		t.Error("expected error for nil updater")
	}

	u = &Updater{}
	err = u.MergeTableCellsVertical(0, 1, 2, 1)
	if err == nil {
		t.Error("expected error for tableIndex < 1")
	}
	err = u.MergeTableCellsVertical(1, 1, 2, 0)
	if err == nil {
		t.Error("expected error for col < 1")
	}
	err = u.MergeTableCellsVertical(1, 0, 2, 1)
	if err == nil {
		t.Error("expected error for startRow < 1")
	}
	err = u.MergeTableCellsVertical(1, 2, 1, 1)
	if err == nil {
		t.Error("expected error for endRow <= startRow")
	}
}

// --- comment.go additional coverage ---

func TestInsertComment_EmptyText(t *testing.T) {
	u := &Updater{}
	err := u.InsertComment(CommentOptions{Anchor: "text"})
	if err == nil || !strings.Contains(err.Error(), "comment text cannot be empty") {
		t.Error("expected empty text error")
	}
}

func TestInsertComment_EmptyAnchor(t *testing.T) {
	u := &Updater{}
	err := u.InsertComment(CommentOptions{Text: "comment"})
	if err == nil || !strings.Contains(err.Error(), "anchor text cannot be empty") {
		t.Error("expected empty anchor error")
	}
}

func TestGetComments_Validation(t *testing.T) {
	var u *Updater
	_, err := u.GetComments()
	if err == nil {
		t.Error("expected error for nil updater")
	}
}
