# docx-viewer-library

`docx-viewer-library` is a TypeScript library for parsing DOCX files and rendering document content as HTML, with specific focus on tracked changes and revision metadata.

The current codebase includes:

- A TypeScript parsing/rendering core under `src/`
- A FastAPI demo application under `demo/`
- Initial support for tracked insertions and deletions, revision authors, paragraph styles, and Word-like table rendering

## Features

- Parse `word/document.xml` from a `.docx` file
- Extract tracked changes such as `w:ins` and `w:del`
- Preserve revision metadata including author, date, and revision id
- Render HTML for paragraphs, inline revisions, and tables
- Map paragraph styles such as title, subtitle, and headings
- Demo viewer with:
  - tracked changes vs applied changes toggle
  - revision gutter markers
  - hover details for revisions

## Project Status

This project is functional but not complete.

Current implementation is strongest for:

- tracked changes inspection
- WordprocessingML parsing experiments
- Word-like rendering for structured technical documents
- CR-style DOCX documents with tables and revision markup

Known limitations:

- Word layout fidelity is still approximate
- Table border conflict resolution is not fully Word-compatible
- Numbering, lists, themes, and some advanced style inheritance are incomplete
- Move tracking and formatting-change revisions need more work

## Installation

### Library prerequisites

- Node.js 18+ recommended
- npm, pnpm, or yarn

### Install dependencies

```bash
npm install
```

### Build

```bash
npm run build
```

### Type check

```bash
npm run check
```

## Usage

```ts
import { parseDocx, renderDocumentAsHtml } from "docx-viewer-library";

const documentModel = await parseDocx(buffer);
const html = renderDocumentAsHtml(documentModel);
```

The parser returns an intermediate document model. The HTML renderer then converts that model into browser-friendly markup.

## Demo Application

The repository includes a FastAPI demo so you can upload a `.docx` file and inspect the rendered result in a browser.

### Demo prerequisites

- Python 3.10+ recommended
- A virtual environment activated by the user

### Install demo dependencies

```bash
pip install -r requirements-demo.txt
```

### Run the demo

```bash
uvicorn demo.app:app --reload
```

Then open:

```text
http://127.0.0.1:8000/
```

If you prefer, you can also use:

```bash
bash run_demo.sh
```

That script assumes your Python environment is already activated.

## Repository Layout

```text
src/                  TypeScript parser and renderer
demo/                 FastAPI demo application
requirements-demo.txt Python dependencies for the demo
run_demo.sh           Bash launcher for the demo
```

## Roadmap

- Improve Word-compatible table rendering
- Expand tracked changes coverage beyond insertions and deletions
- Support numbering and list rendering
- Improve style inheritance and theme handling
- Add tests with representative DOCX samples

## License

MIT
