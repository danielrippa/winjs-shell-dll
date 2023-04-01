
  { file-open-dialog } = winjs.load-library 'WinjsShell.dll'

  process.io.stdout file-open-dialog 'document.md', 'md', 'Markdown files (*.md)', 'c:\\', 'Open Markdow File', 1