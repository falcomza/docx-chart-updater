package godocx

import (
	"bytes"
	"os"
	"testing"
)

func TestNewFromReader_NilReader(t *testing.T) {
	_, err := NewFromReader(nil)
	if err == nil {
		t.Error("expected error for nil reader")
	}
}

func TestSaveToWriter_NilUpdater(t *testing.T) {
	var u *Updater
	var buf bytes.Buffer
	err := u.SaveToWriter(&buf)
	if err == nil {
		t.Error("expected error for nil updater")
	}
}

func TestSaveToWriter_NilWriter(t *testing.T) {
	// We need a valid Updater for this test; create a minimal temp dir
	tmpDir, err := os.MkdirTemp("", "docx-test-*")
	if err != nil {
		t.Fatalf("create temp dir: %v", err)
	}
	defer os.RemoveAll(tmpDir)

	u := &Updater{tempDir: tmpDir}
	err = u.SaveToWriter(nil)
	if err == nil {
		t.Error("expected error for nil writer")
	}
}

func TestWriteZipFromDir(t *testing.T) {
	// Create a minimal directory structure
	tmpDir, err := os.MkdirTemp("", "zip-test-*")
	if err != nil {
		t.Fatalf("create temp dir: %v", err)
	}
	defer os.RemoveAll(tmpDir)

	// Write a test file
	if err := os.WriteFile(tmpDir+"/test.txt", []byte("hello"), 0o644); err != nil {
		t.Fatalf("write test file: %v", err)
	}

	var buf bytes.Buffer
	if err := writeZipFromDir(tmpDir, &buf); err != nil {
		t.Fatalf("writeZipFromDir: %v", err)
	}

	if buf.Len() == 0 {
		t.Error("expected non-empty zip output")
	}
}
