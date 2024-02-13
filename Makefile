PACKAGE_NAME := pyexcelize

ifeq ($(OS),Windows_NT)
    # Windows
    SHARED_LIBRARY := pyexcelize.dll
else
    # Linux
    SHARED_LIBRARY := pyexcelize.so
endif

GO_SOURCE := pyexcelize.go
CGO_FLAGS := -buildmode=c-shared
TARGET_FOLDER := ./pyexcelize
CLEAN_CMD := rm -f $(SHARED_LIBRARY)

all: build

build: $(SHARED_LIBRARY)

$(SHARED_LIBRARY): $(GO_SOURCE)
	@echo "Building shared library..."
	go build $(CGO_FLAGS) -o $(TARGET_FOLDER)/$(SHARED_LIBRARY)

clean:
	@echo "Cleaning up..."
	$(CLEAN_CMD)
