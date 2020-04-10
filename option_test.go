package xlsx_test

import (
	"errors"
	"testing"

	"github.com/bingoohuang/xlsx"
	"github.com/stretchr/testify/assert"
)

type errReader int

func (errReader) Read(p []byte) (n int, err error) {
	return 0, errors.New("test error")
}

func TestWithExcel(t *testing.T) {
	assert.Nil(t, xlsx.WithExcel(errReader(0)))
	assert.Nil(t, xlsx.WithExcel("README.md"))
	assert.Nil(t, xlsx.WithTemplate("README.md"))
	assert.Nil(t, xlsx.WithTemplate(t))
}
