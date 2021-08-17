package upload

import (
	"bytes"
	"context"
	"io"
	"mime"
	"mime/multipart"
	"net/http"
	"os"
	"path/filepath"
)

// Upload uploads to uri with the file specified  in path.
func Upload(ctx context.Context, uri, path, fileFieldName string,
	extraParams map[string]string) (*bytes.Buffer, string, error) {
	request, err := NewUploadRequest(ctx, uri, path, fileFieldName, extraParams)
	if err != nil {
		return nil, "", err
	}

	client := &http.Client{}
	resp, err := client.Do(request)
	if err != nil {
		return nil, "", err
	}

	defer resp.Body.Close()

	body := &bytes.Buffer{}
	if _, err := body.ReadFrom(resp.Body); err != nil {
		return nil, "", err
	}

	fn := DecodeDownloadFilename(resp)

	return body, fn, nil
}

// NewUploadRequest Creates a new file Upload http request with optional extra params.
func NewUploadRequest(ctx context.Context, uri string, path, fileFieldName string,
	extraParams map[string]string) (*http.Request, error) {
	file, err := os.Open(path)
	if err != nil {
		return nil, err
	}

	defer file.Close()

	body := &bytes.Buffer{}
	writer := multipart.NewWriter(body)

	if fileFieldName == "" {
		fileFieldName = "file"
	}

	part, err := writer.CreateFormFile(fileFieldName, filepath.Base(path))
	if err != nil {
		return nil, err
	}

	if _, err = io.Copy(part, file); err != nil {
		return nil, err
	}

	for key, val := range extraParams {
		_ = writer.WriteField(key, val)
	}

	if err := writer.Close(); err != nil {
		return nil, err
	}

	req, err := http.NewRequestWithContext(ctx, "POST", uri, body)
	if err != nil {
		return nil, err
	}

	req.Header.Set("Content-Type", writer.FormDataContentType())

	return req, nil
}

// DecodeDownloadFilename decodes the filename from the http response or empty.
func DecodeDownloadFilename(res *http.Response) string {
	// decode w.Header().Set("Content-Disposition", "attachment; filename=WHATEVER_YOU_WANT")
	if cd := res.Header.Get("Content-Disposition"); cd != "" {
		if _, params, err := mime.ParseMediaType(cd); err == nil {
			return params["filename"]
		}
	}

	return ""
}
