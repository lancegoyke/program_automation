# Program Automation

This is a Python script I run every month to copy training programs to multiple clients' spreadsheets utilizing the Google Sheets API.

## Setup

This project uses [uv](https://docs.astral.sh/uv/) for dependency management.

```sh
uv sync
```

## Run

```sh
uv run autoprog.py
```

## Manage dependencies

```sh
uv add <package>
uv remove <package>
```
