PACKAGE_NAME := pyfastexcel

ifeq ($(OS),Windows_NT)
    # Windows
    SHARED_LIBRARY := pyfastexcel.dll
else
    UNAME_S := $(shell uname -s)
    ifeq ($(UNAME_S),Linux)
        # Linux
        SHARED_LIBRARY := pyfastexcel.so
    else ifeq ($(UNAME_S),Darwin)
        # macOS
        SHARED_LIBRARY := pyfastexcel.dylib
    endif
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

test:
	@echo "Running tests with pytest..."
	pytest -s -v --cov --cov-report=term --cov-report=html
