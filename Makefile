PACKAGE_NAME := pyfastexcel
TARGET_FOLDER := ./pyfastexcel

ifeq ($(OS),Windows_NT)
    # Windows
    SHARED_LIBRARY_NAME := pyfastexcel.dll
else
    UNAME_S := $(shell uname -s)
    ifeq ($(UNAME_S),Linux)
        # Linux
        SHARED_LIBRARY_NAME := pyfastexcel.so
    else ifeq ($(UNAME_S),Darwin)
        # macOS
        SHARED_LIBRARY_NAME := pyfastexcel.dylib
    endif
endif

GO_SOURCE := pyfastexcel.go
GO_SOURCES := $(GO_SOURCE) $(filter-out %_test.go,$(wildcard pyfastexcel/core/*.go))
CGO_FLAGS := -buildmode=c-shared
SHARED_LIBRARY := $(TARGET_FOLDER)/$(SHARED_LIBRARY_NAME)
CLEAN_CMD := rm -f $(SHARED_LIBRARY)

all: build

build: $(SHARED_LIBRARY)

$(SHARED_LIBRARY): $(GO_SOURCES) go.mod go.sum
	@echo "Building shared library..."
	go build $(CGO_FLAGS) -o $(SHARED_LIBRARY)

clean:
	@echo "Cleaning up..."
	$(CLEAN_CMD)

test:
	@echo "Running tests with pytest..."
	uv run pytest -s -v --cov --cov-report=term --cov-report=html

install-dev:
	uv sync --dev

install-docs:
	uv sync --group docs

build-package:
	uv build

lint:
	uv run ruff check .
	uv run ruff format --check .

format:
	uv run ruff check --fix .
	uv run ruff format .

.PHONY: all build clean test install-dev install-docs build-package lint format
