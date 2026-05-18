# Initial Architecture

## Layers

1. `parseDocx`
   Reads the ZIP container and parses `word/document.xml`.
2. `model`
   Defines the stable intermediate model consumed by the rest of the stack.
3. `renderHtml`
   Converts the model into semantic HTML.
4. `styles`
   Provides a minimal stylesheet for visualizing revisions.

## Decisions

- The parser does not render HTML directly.
- Revisions are represented as dedicated nodes with author, date, and id metadata.
- The renderer uses `ins` and `del` because they map well to tracked insertions and deletions in HTML.

## Recommended Next Extensions

- Support `w:moveFrom` and `w:moveTo` with pairing by id.
- Handle formatting and property changes (`w:rPrChange`, `w:pPrChange`).
- Load styles, numbering, tables, headers, and other document structures.
- Add tests with real DOCX samples that cover nested revisions and multiple authors.
