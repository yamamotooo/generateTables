BUILD_DIR := build
SRC       := ./cmd/main

.PHONY: build darwin-arm64 darwin-amd64 windows clean

build:
	go build -o $(BUILD_DIR)/generateTables $(SRC)

darwin-arm64:
	CGO_ENABLED=0 GOOS=darwin GOARCH=arm64 go build -o $(BUILD_DIR)/generateTables_darwin_arm64 $(SRC)

darwin-amd64:
	CGO_ENABLED=0 GOOS=darwin GOARCH=amd64 go build -o $(BUILD_DIR)/generateTables_darwin_amd64 $(SRC)

windows:
	CGO_ENABLED=0 GOOS=windows GOARCH=amd64 go build -o $(BUILD_DIR)/generateTables.exe $(SRC)

clean:
	rm -f $(BUILD_DIR)/generateTables $(BUILD_DIR)/generateTables_darwin_* $(BUILD_DIR)/generateTables.exe $(BUILD_DIR)/output.xml
