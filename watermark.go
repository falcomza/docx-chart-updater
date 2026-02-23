package godocx

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"
)

// WatermarkOptions defines options for text watermarks
type WatermarkOptions struct {
	// Text to display (e.g., "DRAFT", "CONFIDENTIAL")
	Text string

	// FontFamily for the watermark text (default: "Calibri")
	FontFamily string

	// Color as hex code without '#' (default: "C0C0C0" for silver)
	Color string

	// Opacity from 0.0 to 1.0 (default: 0.5)
	Opacity float64

	// Diagonal rotates text at -45 degrees (default: true)
	Diagonal bool
}

// DefaultWatermarkOptions returns watermark options with sensible defaults
func DefaultWatermarkOptions() WatermarkOptions {
	return WatermarkOptions{
		Text:       "DRAFT",
		FontFamily: "Calibri",
		Color:      "C0C0C0",
		Opacity:    0.5,
		Diagonal:   true,
	}
}

// SetTextWatermark adds a text watermark to the document.
// The watermark is inserted into the default header so it appears on every page.
// If a default header already exists, the watermark is injected into it.
func (u *Updater) SetTextWatermark(opts WatermarkOptions) error {
	if u == nil {
		return fmt.Errorf("updater is nil")
	}
	if opts.Text == "" {
		return fmt.Errorf("watermark text cannot be empty")
	}

	// Apply defaults
	if opts.FontFamily == "" {
		opts.FontFamily = "Calibri"
	}
	if opts.Color == "" {
		opts.Color = "C0C0C0"
	}
	if opts.Opacity <= 0 {
		opts.Opacity = 0.5
	}
	if opts.Opacity > 1.0 {
		opts.Opacity = 1.0
	}

	// Normalize color
	if c := normalizeHexColor(opts.Color); c != "" {
		opts.Color = c
	}

	// Generate watermark paragraph XML
	watermarkXML := generateWatermarkShapeXML(opts)

	// Find or create the default header
	headerFile, err := u.findDefaultHeaderFile()
	if err != nil {
		return fmt.Errorf("find default header: %w", err)
	}

	if headerFile != "" {
		// Inject watermark into existing header
		if err := u.injectWatermarkIntoHeader(headerFile, watermarkXML); err != nil {
			return fmt.Errorf("inject watermark: %w", err)
		}
	} else {
		// Create new header with watermark
		if err := u.createWatermarkHeader(watermarkXML); err != nil {
			return fmt.Errorf("create watermark header: %w", err)
		}
	}

	return nil
}

// findDefaultHeaderFile finds the filename of the default header, or "" if none exists.
func (u *Updater) findDefaultHeaderFile() (string, error) {
	docPath := filepath.Join(u.tempDir, "word", "document.xml")
	raw, err := os.ReadFile(docPath)
	if err != nil {
		return "", fmt.Errorf("read document.xml: %w", err)
	}

	// Find headerReference with type="default"
	refPattern := regexp.MustCompile(`<w:headerReference w:type="default" r:id="([^"]+)"/>`)
	matches := refPattern.FindSubmatch(raw)
	if matches == nil {
		return "", nil
	}

	relID := string(matches[1])

	// Look up the target file in document.xml.rels
	relsPath := filepath.Join(u.tempDir, "word", "_rels", "document.xml.rels")
	relsRaw, err := os.ReadFile(relsPath)
	if err != nil {
		return "", fmt.Errorf("read document.xml.rels: %w", err)
	}

	targetPattern := regexp.MustCompile(fmt.Sprintf(`<Relationship Id="%s"[^>]*Target="([^"]+)"`, regexp.QuoteMeta(relID)))
	targetMatches := targetPattern.FindSubmatch(relsRaw)
	if targetMatches == nil {
		return "", nil
	}

	return string(targetMatches[1]), nil
}

// injectWatermarkIntoHeader adds the watermark paragraph to an existing header file.
func (u *Updater) injectWatermarkIntoHeader(headerFile string, watermarkXML []byte) error {
	headerPath := filepath.Join(u.tempDir, "word", headerFile)
	raw, err := os.ReadFile(headerPath)
	if err != nil {
		return fmt.Errorf("read header %s: %w", headerFile, err)
	}

	content := string(raw)

	// Ensure VML namespaces are present on the root element
	content = ensureVMLNamespaces(content)

	// Find the first '>' that's part of the <w:hdr...> opening tag
	hdrIdx := strings.Index(content, "<w:hdr")
	if hdrIdx == -1 {
		return fmt.Errorf("could not find <w:hdr> element")
	}
	hdrCloseIdx := strings.Index(content[hdrIdx:], ">")
	if hdrCloseIdx == -1 {
		return fmt.Errorf("malformed <w:hdr> element")
	}
	insertPos := hdrIdx + hdrCloseIdx + 1

	updated := content[:insertPos] + "\n" + string(watermarkXML) + content[insertPos:]

	if err := os.WriteFile(headerPath, []byte(updated), 0o644); err != nil {
		return fmt.Errorf("write header: %w", err)
	}

	return nil
}

// createWatermarkHeader creates a new default header file containing the watermark.
func (u *Updater) createWatermarkHeader(watermarkXML []byte) error {
	headerFile := "header1.xml"
	headerPath := filepath.Join(u.tempDir, "word", headerFile)

	// Check if header1.xml already exists (used for something else)
	for i := 1; i <= 10; i++ {
		headerFile = fmt.Sprintf("header%d.xml", i)
		headerPath = filepath.Join(u.tempDir, "word", headerFile)
		if _, err := os.Stat(headerPath); os.IsNotExist(err) {
			break
		}
	}

	// Generate full header XML with watermark
	var buf strings.Builder
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString("\n")
	buf.WriteString(`<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" `)
	buf.WriteString(`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" `)
	buf.WriteString(`xmlns:v="urn:schemas-microsoft-com:vml" `)
	buf.WriteString(`xmlns:o="urn:schemas-microsoft-com:office:office">`)
	buf.WriteString("\n")
	buf.Write(watermarkXML)
	buf.WriteString("\n")
	buf.WriteString(`</w:hdr>`)

	if err := os.WriteFile(headerPath, []byte(buf.String()), 0o644); err != nil {
		return fmt.Errorf("write header: %w", err)
	}

	// Add relationship
	relID, err := u.addHeaderFooterRelationship(headerFile, "header")
	if err != nil {
		return fmt.Errorf("add relationship: %w", err)
	}

	// Update document.xml sectPr with header reference
	if err := u.updateDocumentForHeaderFooter("default", "header", relID, false, false); err != nil {
		return fmt.Errorf("update document: %w", err)
	}

	// Add content type
	if err := u.addHeaderFooterContentType(headerFile, "header"); err != nil {
		return fmt.Errorf("add content type: %w", err)
	}

	return nil
}

// generateWatermarkShapeXML creates the VML shape XML for a text watermark.
func generateWatermarkShapeXML(opts WatermarkOptions) []byte {
	var buf bytes.Buffer

	rotation := ""
	if opts.Diagonal {
		rotation = "rotation:315;"
	}

	buf.WriteString("<w:p>")
	buf.WriteString("<w:pPr><w:pStyle w:val=\"Header\"/></w:pPr>")
	buf.WriteString("<w:r>")
	buf.WriteString("<w:rPr><w:noProof/></w:rPr>")
	buf.WriteString("<w:pict>")

	// VML shapetype for TextPlainText (type 136)
	buf.WriteString(`<v:shapetype id="_x0000_t136" coordsize="21600,21600" `)
	buf.WriteString(`o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">`)
	buf.WriteString(`<v:formulas>`)
	buf.WriteString(`<v:f eqn="sum #0 0 10800"/>`)
	buf.WriteString(`<v:f eqn="prod #0 2 1"/>`)
	buf.WriteString(`<v:f eqn="sum 21600 0 @1"/>`)
	buf.WriteString(`<v:f eqn="sum 0 0 @2"/>`)
	buf.WriteString(`<v:f eqn="sum 21600 0 @3"/>`)
	buf.WriteString(`<v:f eqn="if @0 @3 0"/>`)
	buf.WriteString(`<v:f eqn="if @0 21600 @1"/>`)
	buf.WriteString(`<v:f eqn="if @0 0 @2"/>`)
	buf.WriteString(`<v:f eqn="if @0 @4 21600"/>`)
	buf.WriteString(`<v:f eqn="mid @5 @6"/>`)
	buf.WriteString(`<v:f eqn="mid @8 @5"/>`)
	buf.WriteString(`<v:f eqn="mid @7 @8"/>`)
	buf.WriteString(`<v:f eqn="mid @6 @7"/>`)
	buf.WriteString(`<v:f eqn="sum @6 0 @5"/>`)
	buf.WriteString(`</v:formulas>`)
	buf.WriteString(`<v:path textpathok="t" o:connecttype="custom" `)
	buf.WriteString(`o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" `)
	buf.WriteString(`o:connectangles="270,180,90,0"/>`)
	buf.WriteString(`<v:textpath on="t" fitshape="t"/>`)
	buf.WriteString(`<v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>`)
	buf.WriteString(`<o:lock v:ext="edit" text="t" shapetype="t"/>`)
	buf.WriteString(`</v:shapetype>`)

	// Watermark shape
	buf.WriteString(fmt.Sprintf(`<v:shape id="PowerPlusWaterMarkObject" `+
		`o:spid="_x0000_s2049" type="#_x0000_t136" `+
		`style="position:absolute;margin-left:0;margin-top:0;`+
		`width:468pt;height:117pt;%s`+
		`z-index:-251658752;`+
		`mso-position-horizontal:center;mso-position-horizontal-relative:margin;`+
		`mso-position-vertical:center;mso-position-vertical-relative:margin" `+
		`o:allowincell="f" fillcolor="#%s" stroked="f">`,
		rotation, opts.Color))

	buf.WriteString(fmt.Sprintf(`<v:fill opacity="%.2f"/>`, opts.Opacity))

	buf.WriteString(fmt.Sprintf(`<v:textpath style="font-family:&quot;%s&quot;;font-size:1pt" string="%s"/>`,
		xmlEscape(opts.FontFamily), xmlEscape(opts.Text)))

	buf.WriteString(`</v:shape>`)

	buf.WriteString("</w:pict>")
	buf.WriteString("</w:r>")
	buf.WriteString("</w:p>")

	return buf.Bytes()
}

// ensureVMLNamespaces ensures the VML namespace declarations are present
// in the header XML root element.
func ensureVMLNamespaces(headerXML string) string {
	if !strings.Contains(headerXML, `xmlns:v=`) {
		headerXML = strings.Replace(headerXML,
			`<w:hdr `,
			`<w:hdr xmlns:v="urn:schemas-microsoft-com:vml" `,
			1)
	}
	if !strings.Contains(headerXML, `xmlns:o=`) {
		headerXML = strings.Replace(headerXML,
			`<w:hdr `,
			`<w:hdr xmlns:o="urn:schemas-microsoft-com:office:office" `,
			1)
	}
	return headerXML
}
