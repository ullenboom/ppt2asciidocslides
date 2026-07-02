# PPTX to Markdown converter

This command-line tool converts PowerPoint `.pptx` files into Markdown-based slide documents. It was originally built to migrate large decks of PowerPoint slides to AsciiDoc, but now supports multiple output formats.

The converter reads PPTX files via `org.apache.poi:poi-ooxml` and transforms them into a renderer-neutral DOM. A pluggable `DocumentWriter` then renders the DOM to the selected markup language. Currently two formats are supported:

- **AsciiDoc** (`.adoc`) — the original output format, suited for technical publishing.
- **Marp** (`.md`) — [CommonMark](https://commonmark.org/) with [Marpit](https://marpit.marp.app/) slide directives, suited for generating slide decks with the [Marp ecosystem](https://marp.app/).

## Usage

In the _distribution_ folder is a batch file and a Jar. Either call:

    $ java -jar pptx2markdown.jar Slide.pptx

or

    $ convert Slide.pptx

By default the converter produces AsciiDoc output. Use the `--format` option to choose the output format:

    $ java -jar pptx2markdown.jar --format marp Slide.pptx

## Supported options

| Option        | Description                                          | Default                        |
| ------------- | ---------------------------------------------------- | ------------------------------ |
| `--format`    | Output format: `adoc` (AsciiDoc) or `marp` (Marp)   | `adoc`                         |
| `--target`    | Target filename                                      | Input filename, suffix swapped |
| `--imagefolder` | Folder for extracted images                        | `images`                       |
| `--noimages`  | Skip image extraction                                | _images enabled_               |
| `--help`      | Show help                                            |                                |
