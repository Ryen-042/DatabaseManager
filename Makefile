SOLUTION := DatabaseManager.slnx
APP_PROJECT := src/DatabaseManager.Wpf/DatabaseManager.Wpf.csproj
CONFIG ?= Debug
FRAMEWORK ?= net8.0-windows
RUNTIME ?= win-x64

# ----------------------------------------------------------------------------
# DatabaseManager Makefile
# ----------------------------------------------------------------------------
# What each target does:
# - help: Print available targets and quick usage.
# - restore: Restore NuGet packages for the solution.
# - build: Restore + build the full solution.
# - test: Build + run tests.
# - run: Run the WPF startup project.
# - clean: Clean build outputs (bin/obj) for selected configuration.
# - rebuild: Clean then build.
# - publish: Framework-dependent publish for WPF app.
# - publish-self-contained: Self-contained single-file Windows publish.
# - publish-self-contained-optimized: Optimized self-contained publish (smaller size, WPF-safe).
# - format: Run dotnet format.
#
# Recommended WPF flow from clean to run:
#   make clean CONFIG=Debug
#   make build CONFIG=Debug
#   make run CONFIG=Debug
#
# Recommended validation flow:
#   make clean CONFIG=Debug
#   make test CONFIG=Debug
#   make run CONFIG=Debug
# ----------------------------------------------------------------------------

.PHONY: help restore build test run clean rebuild publish publish-self-contained publish-self-contained-optimized format

help:
	@echo "Available targets:"
	@echo "  make restore                              - Restore NuGet packages"
	@echo "  make build [CONFIG=Debug]                 - Build solution"
	@echo "  make test [CONFIG=Debug]                  - Run tests"
	@echo "  make run                                  - Run WPF app"
	@echo "  make clean                                - Clean build outputs"
	@echo "  make rebuild                              - Clean and build"
	@echo "  make publish [CONFIG=Release]             - Framework-dependent publish"
	@echo "  make publish-self-contained               - Basic self-contained single-file publish"
	@echo "  make publish-self-contained-optimized     - Optimized self-contained publish (smaller size, WPF-safe)"
	@echo "  make format                               - Format code"
	@echo ""
	@echo "WPF clean-to-run flow:"
	@echo "  make clean CONFIG=Debug"
	@echo "  make build CONFIG=Debug"
	@echo "  make run CONFIG=Debug"
	@echo ""
	@echo "WPF validated flow:"
	@echo "  make clean CONFIG=Debug"
	@echo "  make test CONFIG=Debug"
	@echo "  make run CONFIG=Debug"

restore:
	dotnet restore $(SOLUTION)

build: restore
	dotnet build $(SOLUTION) -c $(CONFIG)

test: build
	dotnet test $(SOLUTION) -c $(CONFIG) --no-build

run:
	dotnet run --project $(APP_PROJECT) -c $(CONFIG)

clean:
	dotnet clean $(SOLUTION) -c $(CONFIG)

rebuild: clean build

publish:
	dotnet publish $(APP_PROJECT) -c $(CONFIG) -f $(FRAMEWORK) --self-contained false

publish-self-contained:
	dotnet publish $(APP_PROJECT) \
		-c $(CONFIG) \
		-f $(FRAMEWORK) \
		-r $(RUNTIME) \
		--self-contained true \
		/p:PublishSingleFile=true \
		/p:EnableCompressionInSingleFile=true

publish-self-contained-optimized:
	dotnet publish $(APP_PROJECT) \
		-c Release \
		-f $(FRAMEWORK) \
		-r $(RUNTIME) \
		--self-contained true \
		/p:PublishSingleFile=true \
		/p:EnableCompressionInSingleFile=true \
		/p:DebugType=None \
		/p:DebugSymbols=false

format:
	dotnet format $(SOLUTION)
