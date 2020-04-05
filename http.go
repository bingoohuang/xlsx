package xlsx

import (
	"mime"
	"net/http"

	"github.com/unidoc/unioffice/spreadsheet"
)

type upload struct {
	r           *http.Request
	filenameKey string
}

// WithUpload defines the input excel file for reading.
func WithUpload(r *http.Request, filenameKey string) OptionFn {
	return func(o *Option) { o.httpUpload = &upload{r: r, filenameKey: filenameKey} }
}

func (u *upload) parseUploadFile() (*spreadsheet.Workbook, error) {
	// nolint gomnd
	_ = u.r.ParseMultipartForm(32 << 20) // limit your max input length!

	file, header, err := u.r.FormFile(u.filenameKey)
	if err != nil {
		return nil, err
	}

	defer file.Close()

	return spreadsheet.Read(file, header.Size)
}

// Download downloads the excels file in the http response.
func (x *Xlsx) Download(w http.ResponseWriter, filename string) error {
	h := w.Header().Set

	h("Content-Disposition", createContentDisposition(filename))
	h("Content-Description", "File Transfer")
	h("Content-Type", "application/octet-stream")
	h("Content-Transfer-Encoding", "binary")
	h("Expires", "0")
	h("Cache-Control", "must-revalidate")
	h("Pragma", "public")

	return x.Save(w)
}

func createContentDisposition(filename string) string {
	return mime.FormatMediaType("attachment", map[string]string{"filename": filename})
}
