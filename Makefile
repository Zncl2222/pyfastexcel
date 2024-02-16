PACKAGE_NAME := pyfastexcel

ifeq ($(OS),Windows_NT)
    # Windows
    SHARED_LIBRARY := pyfastexcel.dll
else
    # Linux
    SHARED_LIBRARY := pyfastexcel.so
endif

GO_SOURCE := pyfastexcel.go
CGO_FLAGS := -buildmode=c-shared
TARGET_FOLDER := ./pyfastexcel
CLEAN_CMD := rm -f $(SHARED_LIBRARY)

all: build

build: $(SHARED_LIBRARY)

$(SHARED_LIBRARY): $(GO_SOURCE)
	@echo "Building shared library..."
	go build $(CGO_FLAGS) -o $(TARGET_FOLDER)/$(SHARED_LIBRARY)

clean:
	@echo "Cleaning up..."
	$(CLEAN_CMD)
