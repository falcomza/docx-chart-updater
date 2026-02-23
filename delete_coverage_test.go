package godocx

import (
	"strings"
	"testing"
)

func TestDefaultDeleteOptions(t *testing.T) {
	opts := DefaultDeleteOptions()
	if opts.MatchCase {
		t.Error("expected MatchCase to default to false")
	}
	if opts.WholeWord {
		t.Error("expected WholeWord to default to false")
	}
}

func TestDeleteParagraphsContaining(t *testing.T) {
	tests := []struct {
		name      string
		input     string
		text      string
		opts      DeleteOptions
		wantCount int
		wantKeep  string
		wantGone  string
	}{
		{
			name:      "case insensitive match",
			input:     `<w:body><w:p><w:r><w:t>Hello World</w:t></w:r></w:p><w:p><w:r><w:t>Keep Me</w:t></w:r></w:p></w:body>`,
			text:      "hello",
			opts:      DeleteOptions{MatchCase: false},
			wantCount: 1,
			wantKeep:  "Keep Me",
			wantGone:  "Hello World",
		},
		{
			name:      "case sensitive no match",
			input:     `<w:body><w:p><w:r><w:t>Hello World</w:t></w:r></w:p></w:body>`,
			text:      "hello",
			opts:      DeleteOptions{MatchCase: true},
			wantCount: 0,
			wantKeep:  "Hello World",
		},
		{
			name:      "case sensitive match",
			input:     `<w:body><w:p><w:r><w:t>Hello World</w:t></w:r></w:p><w:p><w:r><w:t>Keep</w:t></w:r></w:p></w:body>`,
			text:      "Hello",
			opts:      DeleteOptions{MatchCase: true},
			wantCount: 1,
			wantGone:  "Hello World",
		},
		{
			name:      "whole word match",
			input:     `<w:body><w:p><w:r><w:t>The cat sat</w:t></w:r></w:p><w:p><w:r><w:t>category</w:t></w:r></w:p></w:body>`,
			text:      "cat",
			opts:      DeleteOptions{WholeWord: true},
			wantCount: 1,
			wantGone:  "The cat sat",
			wantKeep:  "category",
		},
		{
			name:      "whole word case sensitive",
			input:     `<w:body><w:p><w:r><w:t>The Cat sat</w:t></w:r></w:p><w:p><w:r><w:t>the cat ran</w:t></w:r></w:p></w:body>`,
			text:      "cat",
			opts:      DeleteOptions{WholeWord: true, MatchCase: true},
			wantCount: 1,
			wantGone:  "the cat ran",
			wantKeep:  "The Cat sat",
		},
		{
			name:      "multiple matches",
			input:     `<w:body><w:p><w:r><w:t>draft text</w:t></w:r></w:p><w:p><w:r><w:t>another draft</w:t></w:r></w:p><w:p><w:r><w:t>final</w:t></w:r></w:p></w:body>`,
			text:      "draft",
			opts:      DeleteOptions{},
			wantCount: 2,
			wantKeep:  "final",
		},
		{
			name:      "no matches",
			input:     `<w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body>`,
			text:      "nonexistent",
			opts:      DeleteOptions{},
			wantCount: 0,
			wantKeep:  "Hello",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result, count, err := deleteParagraphsContaining([]byte(tt.input), tt.text, tt.opts)
			if err != nil {
				t.Fatalf("unexpected error: %v", err)
			}
			if count != tt.wantCount {
				t.Errorf("got count %d, want %d", count, tt.wantCount)
			}
			rs := string(result)
			if tt.wantKeep != "" && !strings.Contains(rs, tt.wantKeep) {
				t.Errorf("expected to keep %q in result: %s", tt.wantKeep, rs)
			}
			if tt.wantGone != "" && strings.Contains(rs, tt.wantGone) {
				t.Errorf("expected %q to be removed from result: %s", tt.wantGone, rs)
			}
		})
	}
}

func TestDeleteNthTable(t *testing.T) {
	docXML := `<w:body>` +
		`<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Table1</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
		`<w:p><w:r><w:t>Paragraph</w:t></w:r></w:p>` +
		`<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Table2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
		`</w:body>`

	t.Run("delete first table", func(t *testing.T) {
		result, err := deleteNthTable([]byte(docXML), 1)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		rs := string(result)
		if strings.Contains(rs, "Table1") {
			t.Error("expected Table1 to be removed")
		}
		if !strings.Contains(rs, "Table2") {
			t.Error("expected Table2 to be kept")
		}
		if !strings.Contains(rs, "Paragraph") {
			t.Error("expected Paragraph to be kept")
		}
	})

	t.Run("delete second table", func(t *testing.T) {
		result, err := deleteNthTable([]byte(docXML), 2)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		rs := string(result)
		if !strings.Contains(rs, "Table1") {
			t.Error("expected Table1 to be kept")
		}
		if strings.Contains(rs, "Table2") {
			t.Error("expected Table2 to be removed")
		}
	})

	t.Run("table not found", func(t *testing.T) {
		_, err := deleteNthTable([]byte(docXML), 5)
		if err == nil {
			t.Error("expected error for nonexistent table")
		}
	})
}

func TestDeleteNthImage(t *testing.T) {
	docXML := `<w:body>` +
		`<w:p><w:r><w:drawing><wp:inline><a:blip r:embed="rId1"/></wp:inline></w:drawing></w:r></w:p>` +
		`<w:p><w:r><w:t>text between</w:t></w:r></w:p>` +
		`<w:p><w:r><w:drawing><wp:inline><a:blip r:embed="rId2"/></wp:inline></w:drawing></w:r></w:p>` +
		`</w:body>`

	t.Run("delete first image", func(t *testing.T) {
		result, err := deleteNthImage([]byte(docXML), 1)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		rs := string(result)
		if strings.Contains(rs, "rId1") {
			t.Error("expected first image to be removed")
		}
		if !strings.Contains(rs, "rId2") {
			t.Error("expected second image to be kept")
		}
	})

	t.Run("image not found", func(t *testing.T) {
		_, err := deleteNthImage([]byte(docXML), 5)
		if err == nil {
			t.Error("expected error for nonexistent image")
		}
	})
}

func TestDeleteNthChart(t *testing.T) {
	docXML := `<w:body>` +
		`<w:p><w:r><w:drawing><wp:inline><c:chart r:id="rId3"/></wp:inline></w:drawing></w:r></w:p>` +
		`<w:p><w:r><w:t>middle text</w:t></w:r></w:p>` +
		`<w:p><w:r><w:drawing><wp:inline><c:chart r:id="rId4"/></wp:inline></w:drawing></w:r></w:p>` +
		`</w:body>`

	t.Run("delete first chart", func(t *testing.T) {
		result, err := deleteNthChart([]byte(docXML), 1)
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
		rs := string(result)
		if strings.Contains(rs, "rId3") {
			t.Error("expected first chart to be removed")
		}
		if !strings.Contains(rs, "rId4") {
			t.Error("expected second chart to be kept")
		}
	})

	t.Run("chart not found", func(t *testing.T) {
		_, err := deleteNthChart([]byte(docXML), 5)
		if err == nil {
			t.Error("expected error for nonexistent chart")
		}
	})
}

func TestDeleteParagraphs_Validation(t *testing.T) {
	var u *Updater
	_, err := u.DeleteParagraphs("test", DeleteOptions{})
	if err == nil {
		t.Error("expected error for nil updater")
	}

	u = &Updater{}
	_, err = u.DeleteParagraphs("", DeleteOptions{})
	if err == nil {
		t.Error("expected error for empty text")
	}
}

func TestDeleteTable_Validation(t *testing.T) {
	var u *Updater
	err := u.DeleteTable(1)
	if err == nil {
		t.Error("expected error for nil updater")
	}

	u = &Updater{}
	err = u.DeleteTable(0)
	if err == nil {
		t.Error("expected error for index < 1")
	}
}

func TestDeleteImage_Validation(t *testing.T) {
	var u *Updater
	err := u.DeleteImage(1)
	if err == nil {
		t.Error("expected error for nil updater")
	}

	u = &Updater{}
	err = u.DeleteImage(0)
	if err == nil {
		t.Error("expected error for index < 1")
	}
}

func TestDeleteChart_Validation(t *testing.T) {
	var u *Updater
	err := u.DeleteChart(1)
	if err == nil {
		t.Error("expected error for nil updater")
	}

	u = &Updater{}
	err = u.DeleteChart(0)
	if err == nil {
		t.Error("expected error for index < 1")
	}
}

func TestGetTableCount_Validation(t *testing.T) {
	var u *Updater
	_, err := u.GetTableCount()
	if err == nil {
		t.Error("expected error for nil updater")
	}
}

func TestGetParagraphCount_Validation(t *testing.T) {
	var u *Updater
	_, err := u.GetParagraphCount()
	if err == nil {
		t.Error("expected error for nil updater")
	}
}

func TestGetImageCount_Validation(t *testing.T) {
	var u *Updater
	_, err := u.GetImageCount()
	if err == nil {
		t.Error("expected error for nil updater")
	}
}
