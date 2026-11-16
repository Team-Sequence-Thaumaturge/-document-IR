FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PYTHONIOENCODING=utf-8

WORKDIR /app

COPY pyproject.toml README.md LICENSE ./
COPY *.py ./

RUN pip install --no-cache-dir .

ENTRYPOINT ["sadt-mcp"]
