## Simple Google Go (golang) library for replace text in microsoft word (.docx) file

The following constitutes the bare minimum required to replace tex in DOCX document.
#+BEGIN_SRC go 

import (
	"github.com/nguyenthenguyen/docx"
)

func main() {
	r, err := docx.ReadDocxFile("./template.docx")
	if err != nil {
		panic(err)
	}
	r.Replace("<old text>", "new text", -1)
	r.WriteToFile("./new_template.docx")
	r.Close()
}

#+END_SRC