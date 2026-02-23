package godocx

import (
	"strings"
	"testing"
)

func TestParseScatterXValues(t *testing.T) {
	tests := []struct {
		name       string
		categories []string
		wantLen    int
		wantErr    bool
	}{
		{
			name:       "valid integers",
			categories: []string{"1", "2", "3"},
			wantLen:    3,
		},
		{
			name:       "valid floats",
			categories: []string{"1.5", "2.7", "3.14"},
			wantLen:    3,
		},
		{
			name:       "with whitespace",
			categories: []string{" 1.5 ", "  2.0  "},
			wantLen:    2,
		},
		{
			name:       "non-numeric",
			categories: []string{"abc", "def"},
			wantErr:    true,
		},
		{
			name:       "empty string",
			categories: []string{""},
			wantErr:    true,
		},
		{
			name:       "mixed numeric and non-numeric",
			categories: []string{"1.0", "abc"},
			wantErr:    true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result, err := parseScatterXValues(tt.categories)
			if tt.wantErr {
				if err == nil {
					t.Error("expected error")
				}
				return
			}
			if err != nil {
				t.Fatalf("unexpected error: %v", err)
			}
			if len(result) != tt.wantLen {
				t.Errorf("got len %d, want %d", len(result), tt.wantLen)
			}
		})
	}
}

func TestParseScatterXValues_Values(t *testing.T) {
	result, err := parseScatterXValues([]string{"1.5", "2.5", "3.5"})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if result[0] != 1.5 || result[1] != 2.5 || result[2] != 3.5 {
		t.Errorf("unexpected values: %v", result)
	}
}

func TestUpdateChartTitle(t *testing.T) {
	tests := []struct {
		name     string
		content  string
		title    string
		prefix   string
		expected string
	}{
		{
			name:     "update existing title with c: prefix",
			content:  `<c:chart><c:title><c:tx><c:rich><c:t>Old Title</c:t></c:rich></c:tx></c:title></c:chart>`,
			title:    "New Title",
			prefix:   "c:",
			expected: "New Title",
		},
		{
			name:     "no title element",
			content:  `<c:chart><c:plotArea/></c:chart>`,
			title:    "New Title",
			prefix:   "c:",
			expected: "<c:chart><c:plotArea/></c:chart>",
		},
		{
			name:     "no prefix",
			content:  `<chart><title><tx><rich><t>Old</t></rich></tx></title></chart>`,
			title:    "Updated",
			prefix:   "",
			expected: "Updated",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := updateChartTitle(tt.content, tt.title, tt.prefix)
			if !strings.Contains(result, tt.expected) {
				t.Errorf("expected %q in result: %s", tt.expected, result)
			}
		})
	}
}

func TestUpdateAxisTitles(t *testing.T) {
	content := `<c:plotArea>` +
		`<c:catAx><c:title><c:tx><c:rich><c:t>Old Cat</c:t></c:rich></c:tx></c:title></c:catAx>` +
		`<c:valAx><c:title><c:tx><c:rich><c:t>Old Val</c:t></c:rich></c:tx></c:title></c:valAx>` +
		`</c:plotArea>`

	result := updateAxisTitles(content, "New Category", "New Value", "c:")

	if !strings.Contains(result, "New Category") {
		t.Errorf("expected new category title in: %s", result)
	}
	if !strings.Contains(result, "New Value") {
		t.Errorf("expected new value title in: %s", result)
	}
}

func TestUpdateAxisTitles_OnlyCat(t *testing.T) {
	content := `<c:catAx><c:title><c:tx><c:rich><c:t>Old</c:t></c:rich></c:tx></c:title></c:catAx>`

	result := updateAxisTitles(content, "NewCat", "", "c:")

	if !strings.Contains(result, "NewCat") {
		t.Error("expected NewCat in result")
	}
}

func TestUpdateAxisTitles_OnlyVal(t *testing.T) {
	content := `<c:valAx><c:title><c:tx><c:rich><c:t>Old</c:t></c:rich></c:tx></c:title></c:valAx>`

	result := updateAxisTitles(content, "", "NewVal", "c:")

	if !strings.Contains(result, "NewVal") {
		t.Error("expected NewVal in result")
	}
}

func TestUpdateFirstAxisTitle(t *testing.T) {
	tests := []struct {
		name     string
		content  string
		title    string
		prefix   string
		axis     string
		expected string
	}{
		{
			name:     "update catAx title",
			content:  `<c:catAx><c:title><c:tx><c:rich><c:t>Old</c:t></c:rich></c:tx></c:title></c:catAx>`,
			title:    "Categories",
			prefix:   "c:",
			axis:     "catAx",
			expected: "Categories",
		},
		{
			name:     "update valAx title",
			content:  `<c:valAx><c:title><c:tx><c:rich><c:t>Old</c:t></c:rich></c:tx></c:title></c:valAx>`,
			title:    "Values",
			prefix:   "c:",
			axis:     "valAx",
			expected: "Values",
		},
		{
			name:     "axis not found",
			content:  `<c:plotArea><c:lineChart/></c:plotArea>`,
			title:    "Title",
			prefix:   "c:",
			axis:     "catAx",
			expected: "<c:plotArea><c:lineChart/></c:plotArea>",
		},
		{
			name:     "no title in axis",
			content:  `<c:catAx><c:scaling/></c:catAx>`,
			title:    "Title",
			prefix:   "c:",
			axis:     "catAx",
			expected: `<c:catAx><c:scaling/></c:catAx>`,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := updateFirstAxisTitle(tt.content, tt.title, tt.prefix, tt.axis)
			if !strings.Contains(result, tt.expected) {
				t.Errorf("expected %q in result: %s", tt.expected, result)
			}
		})
	}
}

func TestDetectNamespacePrefix(t *testing.T) {
	if p := detectNamespacePrefix("<c:chart>"); p != "c:" {
		t.Errorf("expected 'c:', got %q", p)
	}
	if p := detectNamespacePrefix("<chart>"); p != "" {
		t.Errorf("expected empty prefix, got %q", p)
	}
}

func TestEnsureXMLDeclarationNewline(t *testing.T) {
	tests := []struct {
		name  string
		input string
		want  string
	}{
		{
			name:  "already has newline",
			input: "<?xml version=\"1.0\"?>\n<root/>",
			want:  "<?xml version=\"1.0\"?>\n<root/>",
		},
		{
			name:  "missing newline",
			input: "<?xml version=\"1.0\"?><root/>",
			want:  "<?xml version=\"1.0\"?>\n<root/>",
		},
		{
			name:  "no xml declaration",
			input: "<root/>",
			want:  "<root/>",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := string(ensureXMLDeclarationNewline([]byte(tt.input)))
			if got != tt.want {
				t.Errorf("got %q, want %q", got, tt.want)
			}
		})
	}
}
