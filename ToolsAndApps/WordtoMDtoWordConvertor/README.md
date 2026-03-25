# Pandoc Converter

A lightweight desktop application for converting documents between Word format (.docx) and Markdown format (.md), powered by [Pandoc](https://pandoc.org).

## What it does

- Select a `.docx` or `.md` file and the application detects the format and sets the conversion direction automatically
- Convert with a single button — no command line required
- Choose or create an output folder for each conversion
- Extracts images from Word documents into an `images` subfolder and links them in the converted Markdown file automatically

## Requirements

- [Python 3.7+](https://www.python.org/downloads/) — uses only the standard library (no pip installs needed)
- [Pandoc](https://pandoc.org/installing.html) — must be installed and available on the system PATH

## Getting started

See **[Set Up and Use the Pandoc Converter](./Set%20Up%20and%20Use%20the%20Pandoc%20Converter.md)** for step-by-step instructions covering installation and how to use the application.

## Running the application

Double-click **PandocConverter.pyw** to open the application. No console window appears.

If you prefer the command line:

```bash
python PandocConverter.pyw
```

## File structure

```
Pandoc Converter/
├── PandocConverter.pyw                          # Application
├── Set Up and Use the Pandoc Converter.md       # Setup and user guide
└── README.md                                    # This file
```
