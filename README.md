# Program Automation

This is a Python script I run every month to copy training programs to multiple clients' spreadsheets utilizing the Google Sheets API.

## Setup

This project uses [uv](https://docs.astral.sh/uv/) for dependency management.

```sh
uv sync
```

## Run

Check the Google connection and preview the run first. It's read-only and copies nothing:

```sh
uv run autoprog.py --check
```

Then copy the program (set `PROGRAM_NAME` in `autoprog.py` first):

```sh
uv run autoprog.py
```

If `token.json` is missing or revoked, either command opens a browser to sign in to Google again.

## Manage dependencies

```sh
uv add <package>
uv remove <package>
```
