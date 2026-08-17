/* MarkDexam — docxBuilder.js
   Converts a restricted "exam markdown" dialect into a styled .docx
   using the `docx` library (loaded globally as window.docx via CDN).

   Syntax:
     #  text        -> Main Title      (Times New Roman, 16pt, bold, navy, center)
     ## text        -> Subtitle        (Times New Roman, 13pt, bold, soft-navy, center)
     ### text       -> Class/Marks line(Arial, 12pt, bold, black, left)
     ---            -> horizontal divider
     > text         -> Section banner  (solid navy box, white bold text)
     (anything else)-> normal body paragraph, plain text
                       - "(a)" "(b)" "(i)" "(iv)" "(I)" "(3)" etc. auto-bold
                       - "[2 Marks]" "[1 x 4 = 4 Marks]" etc. auto-bold + colored #4a607a
                       - lines indented with 2+ spaces / a tab get a deeper
                         left indent (sub-option style)
     (blank line)   -> small vertical spacer
*/

(function (global) {
  const NAVY = "1b365d";
  const NAVY_SOFT = "4a607a";
  const RULE_GREY = "A0A0A0";
  const CONTENT_WIDTH = 10466; // A4 minus 0.5in margins on both sides

  // matches (a) (B) (iv) (III) (12) style short parenthetical markers
  const BOLD_TOKEN_RE = /\((?:[a-zA-Z]{1,5}|\d{1,3})\)/g;
  // matches [2 Marks] [1 x 4 = 4 Marks] [3 marks] style bracketed mark tags
  const MARKS_TOKEN_RE = /\[[^\[\]]*\bMarks?\b[^\[\]]*\]/gi;

  // Finds all auto-style tokens in text, sorted, non-overlapping (first match wins).
  function findTokens(text) {
    const tokens = [];
    let m;
    BOLD_TOKEN_RE.lastIndex = 0;
    while ((m = BOLD_TOKEN_RE.exec(text)) !== null) {
      tokens.push({ start: m.index, end: m.index + m[0].length, text: m[0], type: "option" });
    }
    MARKS_TOKEN_RE.lastIndex = 0;
    while ((m = MARKS_TOKEN_RE.exec(text)) !== null) {
      tokens.push({ start: m.index, end: m.index + m[0].length, text: m[0], type: "marks" });
    }
    tokens.sort((a, b) => a.start - b.start);
    const filtered = [];
    let lastEnd = -1;
    for (const t of tokens) {
      if (t.start >= lastEnd) {
        filtered.push(t);
        lastEnd = t.end;
      }
    }
    return filtered;
  }

  function parseInlineRuns(text, extra) {
    extra = extra || {};
    const runs = [];
    let lastIndex = 0;
    const tokens = findTokens(text);
    for (const t of tokens) {
      if (t.start > lastIndex) {
        runs.push(new docx.TextRun(Object.assign({ text: text.slice(lastIndex, t.start) }, extra)));
      }
      if (t.type === "marks") {
        runs.push(new docx.TextRun(Object.assign({ text: t.text, bold: true, color: NAVY_SOFT }, extra)));
      } else {
        runs.push(new docx.TextRun(Object.assign({ text: t.text, bold: true }, extra)));
      }
      lastIndex = t.end;
    }
    if (lastIndex < text.length) {
      runs.push(new docx.TextRun(Object.assign({ text: text.slice(lastIndex) }, extra)));
    }
    if (runs.length === 0) runs.push(new docx.TextRun(Object.assign({ text: "" }, extra)));
    return runs;
  }

  function titleParagraph(text) {
    return new docx.Paragraph({
      alignment: docx.AlignmentType.CENTER,
      spacing: { after: 60, line: 180, lineRule: "auto" },
      children: [new docx.TextRun({ text, bold: true, color: NAVY, size: 32, font: "Times New Roman" })],
    });
  }

  function subtitleParagraph(text) {
    return new docx.Paragraph({
      alignment: docx.AlignmentType.CENTER,
      spacing: { after: 60, line: 180, lineRule: "auto" },
      children: [new docx.TextRun({ text, bold: true, color: NAVY_SOFT, size: 26, font: "Times New Roman" })],
    });
  }

  function classLineParagraph(text) {
    return new docx.Paragraph({
      spacing: { after: 60, line: 180, lineRule: "auto" },
      children: [new docx.TextRun({ text, bold: true, size: 24, font: "Arial" })],
    });
  }

  function hrParagraph() {
    return new docx.Paragraph({
      spacing: { after: 200 },
      border: { bottom: { style: docx.BorderStyle.SINGLE, size: 6, color: RULE_GREY, space: 1 } },
      children: [new docx.TextRun({ text: "" })],
    });
  }

  function bannerTable(text) {
    return new docx.Table({
      width: { size: CONTENT_WIDTH, type: docx.WidthType.DXA },
      alignment: docx.AlignmentType.CENTER,
      columnWidths: [CONTENT_WIDTH],
      rows: [
        new docx.TableRow({
          children: [
            new docx.TableCell({
              width: { size: CONTENT_WIDTH, type: docx.WidthType.DXA },
              shading: { fill: NAVY, type: docx.ShadingType.CLEAR, color: "auto" },
              margins: { top: 120, bottom: 120, left: 180, right: 180 },
              children: [
                new docx.Paragraph({
                  spacing: { after: 0 },
                  children: [new docx.TextRun({ text, bold: true, color: "ffffff", size: 22 })],
                }),
              ],
            }),
          ],
        }),
      ],
    });
  }

  function bodyParagraph(text, indent) {
    return new docx.Paragraph({
      indent: { left: indent },
      spacing: { after: 100 },
      children: parseInlineRuns(text),
    });
  }

  function spacerParagraph() {
    return new docx.Paragraph({ spacing: { after: 100 }, children: [new docx.TextRun({ text: "" })] });
  }

  function markdownToDoc(markdown) {
    const lines = markdown.replace(/\r\n/g, "\n").split("\n");
    const children = [];

    for (const raw of lines) {
      const trimmed = raw.trim();

      if (trimmed === "") {
        children.push(spacerParagraph());
        continue;
      }
      if (/^-{3,}$/.test(trimmed)) {
        children.push(hrParagraph());
        continue;
      }
      if (trimmed.startsWith("### ")) {
        children.push(classLineParagraph(trimmed.slice(4)));
        continue;
      }
      if (trimmed.startsWith("## ")) {
        children.push(subtitleParagraph(trimmed.slice(3)));
        continue;
      }
      if (trimmed.startsWith("# ")) {
        children.push(titleParagraph(trimmed.slice(2)));
        continue;
      }
      if (trimmed.startsWith("> ")) {
        children.push(bannerTable(trimmed.slice(2)));
        children.push(spacerParagraph());
        continue;
      }

      const leadingSpaces = (raw.match(/^(\s*)/) || ["", ""])[1];
      const nested = leadingSpaces.length >= 2 || raw.startsWith("\t");
      children.push(bodyParagraph(trimmed, nested ? 432 : 288));
    }

    return new docx.Document({
      sections: [
        {
          properties: {
            page: {
              size: { width: 11906, height: 16838 }, // A4 (twips): 210mm x 297mm
              margin: { top: 720, bottom: 720, left: 720, right: 720 }, // 0.5in all sides, per Google Docs
            },
          },
          children,
        },
      ],
    });
  }

  global.MarkDexam = { markdownToDoc, parseInlineRuns, BOLD_TOKEN_RE };
})(window);
