package xlsx

import (
	"mime"
	"net/http"

	"github.com/sirupsen/logrus"

	"github.com/unidoc/unioffice/spreadsheet"
)

// WithUpload defines the input excel file for reading.
func WithUpload(r *http.Request, filenameKey string) OptionFn {
	wb, err := parseUploadFile(r, filenameKey)
	if err != nil {
		logrus.Warnf("failed to open template excel %v", err)
		return nil
	}

	return func(o *Option) { o.Workbook = wb }
}

// nolint:gomnd
func parseUploadFile(r *http.Request, filenameKey string) (*spreadsheet.Workbook, error) {
	_ = r.ParseMultipartForm(32 << 20) // limit your max input length!

	file, header, err := r.FormFile(filenameKey)
	if err != nil {
		return nil, err
	}

	defer file.Close()

	return spreadsheet.Read(file, header.Size)
}

// Download downloads the excels file in the http response.
func (x *Xlsx) Download(w http.ResponseWriter, filename string) error {
	h := w.Header().Set

	h("Content-Disposition", mime.FormatMediaType("attachment", map[string]string{"filename": filename}))
	h("Content-Description", "File Transfer")
	h("Content-Type", "application/octet-stream")
	h("Content-Transfer-Encoding", "binary")
	h("Expires", "0")
	h("Cache-Control", "must-revalidate")
	h("Pragma", "public")

	return x.Save(w)
}
