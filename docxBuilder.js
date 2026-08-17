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
                       - **text** -> bold, *text* -> italic (standard markdown)
                       - lines indented with 2+ spaces / a tab get a deeper
                         left indent (sub-option style)
     (blank line)   -> small vertical spacer
*/

(function (global) {
  const NAVY = "1b365d";
  const NAVY_SOFT = "4a607a";
  const RULE_GREY = "A0A0A0";
  const CONTENT_WIDTH = 10466; // A4 minus 0.5in margins on both sides

  // Single ordered pass: **bold** must be tried before *italic* so a
  // double-star run isn't mis-read as an italic run with a stray star.
  //   group 1/2 -> **bold**        (2 = inner text)
  //   group 3    -> [.. Marks ..]  (kept as-is, incl. brackets)
  //   group 4    -> (a) (iv) (3)   (kept as-is, incl. parens)
  //   group 5/6 -> *italic*        (6 = inner text)
  const INLINE_TOKEN_RE =
    /(\*\*(.+?)\*\*)|(\[[^\[\]]*\bMarks?\b[^\[\]]*\])|(\((?:[a-zA-Z]{1,5}|\d{1,3})\))|(\*(.+?)\*)/g;

  function parseInlineRuns(text, extra) {
    extra = extra || {};
    const runs = [];
    let lastIndex = 0;
    let m;
    INLINE_TOKEN_RE.lastIndex = 0;
    while ((m = INLINE_TOKEN_RE.exec(text)) !== null) {
      if (m.index > lastIndex) {
        runs.push(new docx.TextRun(Object.assign({ text: text.slice(lastIndex, m.index) }, extra)));
      }
      if (m[1] !== undefined) {
        runs.push(new docx.TextRun(Object.assign({ text: m[2], bold: true }, extra)));
      } else if (m[3] !== undefined) {
        runs.push(new docx.TextRun(Object.assign({ text: m[3], bold: true, color: NAVY_SOFT }, extra)));
      } else if (m[4] !== undefined) {
        runs.push(new docx.TextRun(Object.assign({ text: m[4], bold: true }, extra)));
      } else if (m[5] !== undefined) {
        runs.push(new docx.TextRun(Object.assign({ text: m[6], italics: true }, extra)));
      }
      lastIndex = m.index + m[0].length;
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

  global.MarkDexam = { markdownToDoc, parseInlineRuns, INLINE_TOKEN_RE };
})(window);
