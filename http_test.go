package xlsx_test

import (
	"io/ioutil"
	"net/http"
	"net/http/httptest"
	"os"
	"testing"

	"github.com/bingoohuang/gonet/man"
	"github.com/bingoohuang/xlsx"
	"github.com/stretchr/testify/assert"
)

type Poster struct {
	UpDown func(man.URL, man.UploadFile, *man.DownloadFile) `method:"POST"`
}

func TestUpload(t *testing.T) {
	ts := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		x, err := xlsx.New(xlsx.WithUpload(r, "file"))
		if err != nil {
			w.WriteHeader(500)
			return
		}

		if err := x.Download(w, "file.xlsx"); err != nil {
			w.WriteHeader(500)
		}
	}))

	defer ts.Close()

	f, _ := os.Open("testdata/template.xlsx")
	defer f.Close()

	//var buf bytes.Buffer
	//df := &man.DownloadFile{Writer: &buf}
	df := &man.DownloadFile{Writer: ioutil.Discard}
	postMan := func() (p Poster) { man.New(&p); return }()

	postMan.UpDown(man.URL(ts.URL), man.MakeFile("file", "upload.xlsx", f), df)

	// ioutil.WriteFile("testdata/dl.xlsx", buf.Bytes(), 0644)
	assert.Equal(t, "file.xlsx", df.Filename)
}
