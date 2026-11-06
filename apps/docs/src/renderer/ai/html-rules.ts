export const HTML_RULES = [
  'The html tool input is a restricted HTML fragment. Rules:',
  '- Only these tags are allowed: h1 h2 h3 h4 h5 h6 p ul ol li strong em u s a br table thead tbody tr th td pre code blockquote',
  '- Tables: use <th> for the first (header) row; cells contain plain text only (<br> may split lines); nested tables / merged cells are not supported; once inserted the table is protected as a whole, only cell text remains editable',
  '- Use <pre> for code samples (monospace font + shading, line breaks preserved); use <blockquote> for quotations (indent + left bar)',
  '- Use <formula>LaTeX</formula> for math (produces native Word equations): as a top-level block it becomes its own centered paragraph; placed inside <p>/<li>/<h*> text it is an inline formula flowing with the text, e.g. <p>From <formula>E = mc^2</formula> we know…</p>; supports the common subset of \\frac \\sqrt super/subscripts \\sum \\int \\lim matrix environments Greek letters etc., the align environment is not supported; invalid LaTeX fails the whole call — fix and retry',
  '- Do not include <html>/<body>, markdown code fences, or explanatory text',
  "- <sel>…</sel> appears only in the context to mark the user's selection; it is not a tag you may write",
  '- Organize long content into sections with h2/h3 (unless the user only wants a single paragraph)',
  '- Keep the same language as the original document / user instruction, unless translation is requested',
].join('\n')
