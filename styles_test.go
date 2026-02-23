package godocx

import (
	"strings"
	"testing"
)

func TestGenerateStyleXML_ParagraphStyle(t *testing.T) {
	def := StyleDefinition{
		ID:         "CustomHeading",
		Name:       "Custom Heading",
		Type:       StyleTypeParagraph,
		BasedOn:    "Normal",
		NextStyle:  "Normal",
		FontFamily: "Arial",
		FontSize:   28,
		Color:      "2E74B5",
		Bold:       true,
		Alignment:  ParagraphAlignCenter,
		SpaceBefore: 240,
		SpaceAfter:  120,
		KeepNext:   true,
	}

	result := generateStyleXML(def)
	xml := string(result)

	if !strings.Contains(xml, `w:type="paragraph"`) {
		t.Error("expected paragraph type")
	}
	if !strings.Contains(xml, `w:styleId="CustomHeading"`) {
		t.Error("expected style ID")
	}
	if !strings.Contains(xml, `w:val="Custom Heading"`) {
		t.Error("expected style name")
	}
	if !strings.Contains(xml, `<w:basedOn w:val="Normal"/>`) {
		t.Error("expected basedOn")
	}
	if !strings.Contains(xml, `<w:next w:val="Normal"/>`) {
		t.Error("expected next style")
	}
	if !strings.Contains(xml, `w:ascii="Arial"`) {
		t.Error("expected font family")
	}
	if !strings.Contains(xml, `<w:sz w:val="28"/>`) {
		t.Error("expected font size")
	}
	if !strings.Contains(xml, `<w:color w:val="2E74B5"/>`) {
		t.Error("expected color")
	}
	if !strings.Contains(xml, "<w:b/>") {
		t.Error("expected bold")
	}
	if !strings.Contains(xml, `w:val="center"`) {
		t.Error("expected center alignment")
	}
	if !strings.Contains(xml, `w:before="240"`) {
		t.Error("expected space before")
	}
	if !strings.Contains(xml, `w:after="120"`) {
		t.Error("expected space after")
	}
	if !strings.Contains(xml, "<w:keepNext/>") {
		t.Error("expected keepNext")
	}
}

func TestGenerateStyleXML_CharacterStyle(t *testing.T) {
	def := StyleDefinition{
		ID:        "Emphasis",
		Name:      "Strong Emphasis",
		Type:      StyleTypeCharacter,
		Bold:      true,
		Italic:    true,
		Underline: true,
		Color:     "FF0000",
	}

	result := generateStyleXML(def)
	xml := string(result)

	if !strings.Contains(xml, `w:type="character"`) {
		t.Error("expected character type")
	}
	// Character styles should NOT have pPr
	if strings.Contains(xml, "<w:pPr>") {
		t.Error("character style should not have paragraph properties")
	}
	if !strings.Contains(xml, "<w:b/>") {
		t.Error("expected bold")
	}
	if !strings.Contains(xml, "<w:i/>") {
		t.Error("expected italic")
	}
	if !strings.Contains(xml, `<w:u w:val="single"/>`) {
		t.Error("expected underline")
	}
}

func TestGenerateStyleXML_MinimalStyle(t *testing.T) {
	def := StyleDefinition{
		ID:   "Simple",
		Name: "Simple",
		Type: StyleTypeParagraph,
	}

	result := generateStyleXML(def)
	xml := string(result)

	if !strings.Contains(xml, `w:styleId="Simple"`) {
		t.Error("expected style ID")
	}
	// No formatting = no pPr or rPr
	if strings.Contains(xml, "<w:pPr>") {
		t.Error("expected no paragraph properties for minimal style")
	}
	if strings.Contains(xml, "<w:rPr>") {
		t.Error("expected no run properties for minimal style")
	}
}

func TestGenerateStyleXML_OutlineLevel(t *testing.T) {
	def := StyleDefinition{
		ID:           "TOCHeading",
		Name:         "TOC Heading",
		Type:         StyleTypeParagraph,
		OutlineLevel: 1,
		Bold:         true,
		FontSize:     32,
	}

	result := generateStyleXML(def)
	xml := string(result)

	// Outline level 1 maps to w:val="0" (0-indexed)
	if !strings.Contains(xml, `<w:outlineLvl w:val="0"/>`) {
		t.Error("expected outline level 0 (for heading level 1)")
	}
}

func TestGenerateStyleXML_Formatting(t *testing.T) {
	def := StyleDefinition{
		ID:            "Fancy",
		Name:          "Fancy",
		Type:          StyleTypeParagraph,
		Strikethrough: true,
		AllCaps:       true,
		SmallCaps:     false,
		LineSpacing:   480,
		IndentLeft:    720,
		IndentFirst:   360,
		PageBreakBef:  true,
		KeepLines:     true,
	}

	result := generateStyleXML(def)
	xml := string(result)

	if !strings.Contains(xml, "<w:strike/>") {
		t.Error("expected strikethrough")
	}
	if !strings.Contains(xml, "<w:caps/>") {
		t.Error("expected all caps")
	}
	if strings.Contains(xml, "<w:smallCaps/>") {
		t.Error("should not have small caps")
	}
	if !strings.Contains(xml, `w:line="480"`) {
		t.Error("expected line spacing")
	}
	if !strings.Contains(xml, `w:left="720"`) {
		t.Error("expected left indent")
	}
	if !strings.Contains(xml, `w:firstLine="360"`) {
		t.Error("expected first line indent")
	}
	if !strings.Contains(xml, "<w:pageBreakBefore/>") {
		t.Error("expected page break before")
	}
	if !strings.Contains(xml, "<w:keepLines/>") {
		t.Error("expected keep lines")
	}
}

func TestInjectStyle(t *testing.T) {
	stylesXML := []byte(`<?xml version="1.0"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`)

	styleXML := []byte(`<w:style w:type="paragraph" w:styleId="Custom"><w:name w:val="Custom"/></w:style>`)

	result, err := injectStyle(stylesXML, styleXML)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	resultStr := string(result)

	// Should contain both styles
	if !strings.Contains(resultStr, `w:styleId="Normal"`) {
		t.Error("expected original Normal style")
	}
	if !strings.Contains(resultStr, `w:styleId="Custom"`) {
		t.Error("expected injected Custom style")
	}
	// New style should be before </w:styles>
	customIdx := strings.Index(resultStr, `w:styleId="Custom"`)
	closeIdx := strings.Index(resultStr, "</w:styles>")
	if customIdx > closeIdx {
		t.Error("custom style should appear before </w:styles>")
	}
}

func TestGenerateStylesDocument(t *testing.T) {
	styleXML := []byte(`<w:style w:type="paragraph" w:styleId="Test"><w:name w:val="Test"/></w:style>`)

	result := generateStylesDocument(styleXML)
	resultStr := string(result)

	if !strings.Contains(resultStr, `<?xml version`) {
		t.Error("expected XML declaration")
	}
	if !strings.Contains(resultStr, `<w:styles`) {
		t.Error("expected <w:styles> element")
	}
	if !strings.Contains(resultStr, `w:styleId="Test"`) {
		t.Error("expected style content")
	}
	if !strings.Contains(resultStr, `</w:styles>`) {
		t.Error("expected closing tag")
	}
}
