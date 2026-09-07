.PHONY: install test lint demo serve mcp

install:
	python -m pip install -e ".[dev]"

test:
	pytest

lint:
	ruff check .

demo:
	simlab run examples/service_center.yaml

serve:
	simlab serve

mcp:
	simlab-mcp
