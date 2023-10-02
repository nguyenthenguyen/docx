package docx

import (
	"bytes"
	"embed"
	"fmt"
	"io/ioutil"
	"regexp"
	"strings"
	"testing"
)

const testFile = "./TestDocument.docx"
const testFileResult = "./TestDocumentResult.docx"

func loadFile(file string) *Docx {
	r, err := ReadDocxFile(file)
	if err != nil {
		panic(err)
	}

	return r.Editable()
}

func loadFromMemory(file string) *Docx {
	data, _ := ioutil.ReadFile(file)
	r, err := ReadDocxFromMemory(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		panic(err)
	}

	return r.Editable()
}

//go:embed TestDocument.docx
var testFS embed.FS

func loadFromFs(file string) *Docx {
	r, err := ReadDocxFromFS(strings.TrimPrefix(file, "./"), testFS)
	if err != nil {
		panic(err)
	}

	return r.Editable()
}

//Tests that we are able to load a file from a filesystem and do a quick replacement test
func TestReadDocxFromFS(t *testing.T) {
	d := loadFromFs(testFile)
	simpleReplacementTest(d, t)
}

//Tests that we are able to load a file from a memory array of bytes
func TestReadDocxFromMemory(t *testing.T) {
	d := loadFromMemory(testFile)
	simpleReplacementTest(d, t)
}

func simpleReplacementTest(d *Docx, t *testing.T) {
	if d == nil {
		t.Error("Doc should not be nil', got ", d)
	}
	d.Replace("document.", "line1\r\nline2", 1)
	d.WriteToFile(testFileResult)

	d = loadFile(testFileResult)

	if strings.Contains(d.content, "This is a word document") {
		t.Error("Missing 'This is a word document.', got ", d.content)
	}
}

func TestReplace(t *testing.T) {
	tests := []struct {
		name        string
		replaceWith string
		expect      string
	}{
		{"Windows line breaks", "line1\r\nline2", "line1<w:br/>line2"},
		{"Mac line breaks", "line1\rline2", "line1<w:br/>line2"},
		{"Linux line breaks", "line1\nline2", "line1<w:br/>line2"},
		{"Tabs", "line1\tline2", "line1</w:t><w:tab/><w:t>line2"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			d := loadFile(testFile)
			previous := d.content
			d.Replace("document.", tt.replaceWith, 1)
			d.WriteToFile(testFileResult)

			d = loadFile(testFileResult)

			if strings.Contains(d.content, "This is a word document") {
				t.Error("Should not contain 'This is a word document' after replacement, got ", d.content)
			}

			// These are arbitrary start and end points that should include the part that's being replaced, plus a bit extra
			start := 1200
			end := 1450

			if !strings.Contains(d.content, tt.expect) {
				t.Errorf("Expected '%s'\n previous: \n%s\n got: \n%s", tt.expect, extractMiddle(start, end, previous), extractMiddle(start, end, d.content))
			}
		})
	}
}

func TestReplaceRawByCallback(t *testing.T) {
	d := loadFile(testFile)

	d.ReplaceRawByCallback(func(oldContent string) string {
		return regexp.
			MustCompile(`document\.`).
			ReplaceAllStringFunc(oldContent, func(src string) string {
				return "line1\r\nline2"
			})
	})

	d.WriteToFile(testFileResult)
	d = loadFile(testFileResult)
	if strings.Contains(d.content, "word document.") {
		t.Error("Should not contain 'word document.' after replacement, got ", d.content)
	}
}

// To make test failure messages easier to read, trim off the beginning and end of the document to focus on the part expected
// to have changed
func extractMiddle(start, end int, content string) string {
	return fmt.Sprint(content[start:end])
}

func TestReplaceLink(t *testing.T) {
	d := loadFile(testFile)
	d.ReplaceLink("http://example.com/", "https://github.com/nguyenthenguyen/docx", -1)
	d.WriteToFile(testFileResult)

	d = loadFile(testFileResult)

	if strings.Contains(d.links, "http://example.com") {
		t.Error("Missing 'http://example.com', got ", d.links)
	}

	if !strings.Contains(d.links, "https://github.com/nguyenthenguyen/docx") {
		t.Error("Expected 'word', got ", d.links)
	}
}

func TestReplaceHeader(t *testing.T) {
	d := loadFile(testFile)
	d.ReplaceHeader("This is a header.", "newHeader")
	d.WriteToFile(testFileResult)

	d = loadFile(testFileResult)

	headers := d.headers
	found := false
	for _, v := range headers {
		if strings.Contains(v, "This is a header.") {
			t.Error("Missing 'This is a header.', got ", d.content)
		}

		if strings.Contains(v, "newHeader") {
			found = true
		}
	}
	if !found {
		t.Error("Expected 'newHeader', got ", d.headers)
	}
}

func TestReplaceFooter(t *testing.T) {
	d := loadFile(testFile)
	d.ReplaceFooter("This is a footer.", "newFooter")
	d.WriteToFile(testFileResult)

	d = loadFile(testFileResult)

	footers := d.footers
	found := false
	for _, v := range footers {
		if strings.Contains(v, "This is a footer.") {
			t.Error("Missing 'This is a footer.', got ", d.content)
		}

		if strings.Contains(v, "newFooter") {
			found = true
		}
	}
	if !found {
		t.Error("Expected 'newFooter', got ", d.headers)
	}
}

func TestReplaceImage(t *testing.T) {
	d := loadFile(testFile)
	d.ReplaceImage("word/media/image1.png", "./new.png")
	d.WriteToFile(testFileResult)

	d = loadFile(testFileResult)

	images := d.images

	found := false
	for k := range images {
		if k == "word/media/image1.png" {
			found = true
		}
	}
	if !found {
		t.Error("Expected an 'word/media/image1.png', got something else")
	}

}

func TestDocx_ImagesLen(t *testing.T) {
	d := loadFile(testFile)

	if d.ImagesLen() != 1 {
		t.Error("Expected the sum of images is 1, got something else")
	}
}
