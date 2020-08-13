.PHONY: default test
all: default test

gosec:
	go get github.com/securego/gosec/cmd/gosec
sec:
	@gosec ./...
	@echo "[OK] Go security check was completed!"

fmt:
	gofumports -w .
	gofumpt -w .
	gofmt -s -w .
	go mod tidy
	go fmt ./...
	revive .
	goimports -w .
	golangci-lint run --enable-all

init:
	export GOPROXY=https://goproxy.cn

default: install

install: init
	go install -ldflags="-s -w" ./...

test: init
	go test ./...
