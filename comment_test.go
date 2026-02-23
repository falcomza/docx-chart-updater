package godocx

import (
	"strings"
	"testing"
)

func TestGenerateInitialCommentsXML(t *testing.T) {
	result := string(generateInitialCommentsXML())

	if !strings.Contains(result, "<w:comments") {
		t.Error("expected <w:comments root element")
	}
	if !strings.Contains(result, "</w:comments>") {
		t.Error("expected closing </w:comments> tag")
	}
}

func TestGenerateCommentEntry(t *testing.T) {
	opts := CommentOptions{
		Text:     "This needs revision",
		Author:   "John Doe",
		Initials: "JD",
		Anchor:   "test",
	}

	result := string(generateCommentEntry(1, opts))

	if !strings.Contains(result, `w:id="1"`) {
		t.Error("expected comment ID 1")
	}
	if !strings.Contains(result, `w:author="John Doe"`) {
		t.Error("expected author name")
	}
	if !strings.Contains(result, `w:initials="JD"`) {
		t.Error("expected initials")
	}
	if !strings.Contains(result, "This needs revision") {
		t.Error("expected comment text")
	}
	if !strings.Contains(result, "<w:annotationRef/>") {
		t.Error("expected annotation reference")
	}
}

func TestGetNextCommentID(t *testing.T) {
	xml := []byte(`<w:comments>
		<w:comment w:id="1" w:author="A"/>
		<w:comment w:id="3" w:author="B"/>
	</w:comments>`)

	nextID := getNextCommentID(xml)
	if nextID != 4 {
		t.Errorf("expected next ID 4, got %d", nextID)
	}
}

func TestGetNextCommentID_Empty(t *testing.T) {
	xml := []byte(`<w:comments></w:comments>`)

	nextID := getNextCommentID(xml)
	if nextID != 1 {
		t.Errorf("expected next ID 1, got %d", nextID)
	}
}

func TestParseComments(t *testing.T) {
	xml := []byte(`<w:comments>
		<w:comment w:id="1" w:author="Jane" w:date="2024-01-15T10:30:00Z" w:initials="J">
			<w:p><w:r><w:t>Review this section</w:t></w:r></w:p>
		</w:comment>
		<w:comment w:id="2" w:author="Bob" w:date="2024-01-16T14:00:00Z" w:initials="B">
			<w:p><w:r><w:t>Looks good</w:t></w:r></w:p>
		</w:comment>
	</w:comments>`)

	comments := parseComments(xml)
	if len(comments) != 2 {
		t.Fatalf("expected 2 comments, got %d", len(comments))
	}

	if comments[0].ID != 1 {
		t.Errorf("expected comment ID 1, got %d", comments[0].ID)
	}
	if comments[0].Author != "Jane" {
		t.Errorf("expected author Jane, got %s", comments[0].Author)
	}
	if !strings.Contains(comments[0].Text, "Review this section") {
		t.Errorf("expected comment text, got %s", comments[0].Text)
	}

	if comments[1].ID != 2 {
		t.Errorf("expected comment ID 2, got %d", comments[1].ID)
	}
	if comments[1].Author != "Bob" {
		t.Errorf("expected author Bob, got %s", comments[1].Author)
	}
}

func TestInsertComment_Validation(t *testing.T) {
	var u *Updater
	err := u.InsertComment(CommentOptions{Text: "test", Anchor: "test"})
	if err == nil {
		t.Error("expected error for nil updater")
	}
}
