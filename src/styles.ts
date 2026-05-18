export const defaultViewerStyles = `
.docx-viewer {
  color: #1f2937;
  font-family: Georgia, "Times New Roman", serif;
  line-height: 1.6;
}

.docx-viewer ins[data-revision-kind] {
  background: #dcfce7;
  color: #166534;
  text-decoration: none;
}

.docx-viewer del[data-revision-kind] {
  background: #fee2e2;
  color: #991b1b;
}

.docx-viewer ins[data-revision-author]::after,
.docx-viewer del[data-revision-author]::after {
  content: " [" attr(data-revision-author) "]";
  font-size: 0.82em;
  opacity: 0.8;
}
`;
