FROM python:3.12-slim

WORKDIR /app
COPY pyproject.toml README.md ./
COPY src ./src
RUN pip install --no-cache-dir ".[api,mcp]"
COPY examples ./examples

EXPOSE 8000

ENTRYPOINT ["simlab"]
CMD ["run", "examples/service_center.yaml"]
