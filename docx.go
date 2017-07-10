package docx

import (
	"archive/zip"
	"bufio"
	"bytes"
	"encoding/xml"
	"errors"
	"io"
	"io/ioutil"
	"os"
	"regexp"
	"strings"
	"fmt"
)

type ReplaceDocx struct {
	zipReader *zip.ReadCloser
	content   string
	header1   string
	header2   string
	header3   string
}

func (r *ReplaceDocx) Editable() *Docx {
	return &Docx{
		files:   r.zipReader.File,
		content: r.content,
		header1: r.header1,
		header2: r.header2,
		header3: r.header3,
	}
}

func (r *ReplaceDocx) Close() error {
	return r.zipReader.Close()
}

type Docx struct {
	files   []*zip.File
	content string
	header1 string
	header2 string
	header3 string
}

func (d *Docx) Replace(oldString string, newString string, num int) (err error) {
	oldString, err = encode(oldString)
	if err != nil {
		return err
	}
	newString, err = encode(newString)
	if err != nil {
		return err
	}
	d.content = strings.Replace(d.content, oldString, newString, num)

	return nil
}

func (d *Docx) ReplaceHeader(oldString string, newString string, num int) (err error) {
	oldString, err = encode(oldString)
	if err != nil {
		return err
	}
	newString, err = encode(newString)
	if err != nil {
		return err
	}

	d.header1 = strings.Replace(d.content, oldString, newString, num)
	d.header2 = strings.Replace(d.content, oldString, newString, num)
	d.header3 = strings.Replace(d.content, oldString, newString, num)
	return nil
}

func (d *Docx) WriteToFile(path string) (err error) {
	var target *os.File
	target, err = os.Create(path)
	if err != nil {
		return
	}
	defer target.Close()
	err = d.Write(target)
	return
}

func (d *Docx) Write(ioWriter io.Writer) (err error) {
	w := zip.NewWriter(ioWriter)
	for _, file := range d.files {
		var writer io.Writer
		var readCloser io.ReadCloser

		writer, err = w.Create(file.Name)
		if err != nil {
			return err
		}
		readCloser, err = file.Open()
		if err != nil {
			return err
		}
		if file.Name == "word/document.xml" {
			writer.Write([]byte(d.content))
		} else {
			writer.Write(streamToByte(readCloser))
		}
	}
	w.Close()
	return
}

func ReadDocxFile(path string) (*ReplaceDocx, error) {
	reader, err := zip.OpenReader(path)
	if err != nil {
		return nil, err
	}
	content, err := readText(reader.File)
	if err != nil {
		return nil, err
	}

	headers, _ := readHeader(reader.File)
	return &ReplaceDocx{zipReader: reader, content: content, header1: headers[0], header2: headers[1], header3: headers[2]}, nil
}

func readHeader(files []*zip.File) (headerText [3]string, err error) {

	h, err := retrieveHeaderDoc(files)
	if err != nil {
		return [3]string{}, err
	}

	var documentReader io.ReadCloser
	for i, element := range h {
		fmt.Println("len(headerText):", len(headerText))
		documentReader, err = element.Open()
		if err != nil {
			return [3]string{}, err
		}

		text, err := wordDocToString(documentReader)
		if err != nil {
			return [3]string{}, err
		}
		if text != "" {
			headerText[i] = text
		}
	}

	return
}

func readText(files []*zip.File) (text string, err error) {
	var documentFile *zip.File
	documentFile, err = retrieveWordDoc(files)
	if err != nil {
		return text, err
	}
	var documentReader io.ReadCloser
	documentReader, err = documentFile.Open()
	if err != nil {
		return text, err
	}

	text, err = wordDocToString(documentReader)
	return
}

func wordDocToString(reader io.Reader) (string, error) {
	b, err := ioutil.ReadAll(reader)
	if err != nil {
		return "", err
	}
	return string(b), nil
}

func retrieveWordDoc(files []*zip.File) (file *zip.File, err error) {
	for _, f := range files {
		if f.Name == "word/document.xml" {
			file = f
		}
	}
	if file == nil {
		err = errors.New("document.xml file not found")
	}
	return
}

func retrieveHeaderDoc(files []*zip.File) (headers []*zip.File, err error) {
	var h [3]*zip.File

	i := 0
	re := regexp.MustCompile(`word/header[1-3].xml`)
	for _, f := range files {
		if re.MatchString(f.Name) {
			h[i] = f
			i++
		}
	}
	if len(h) == 0 {
		err = errors.New("headers[1-3.xml file not found")
	}
	return
}

func streamToByte(stream io.Reader) []byte {
	buf := new(bytes.Buffer)
	buf.ReadFrom(stream)
	return buf.Bytes()
}

func encode(s string) (string, error) {
	var b bytes.Buffer
	enc := xml.NewEncoder(bufio.NewWriter(&b))
	if err := enc.Encode(s); err != nil {
		return s, err
	}
	return strings.Replace(strings.Replace(b.String(), "<string>", "", 1), "</string>", "", 1), nil
}
