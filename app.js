(function () {
  const editor = document.getElementById("editor");
  const preview = document.getElementById("preview");
  const downloadBtn = document.getElementById("downloadBtn");
  const sampleBtn = document.getElementById("sampleBtn");

  const SAMPLE_MD = `# Vihan Vidyapeeth Pakari Vaina, Ballia
## Half Yearly Examination (2026-27)
### Class: VII                                                                    Time: 3 Hrs
### Subject: Mathematics                                                Maximum Marks: 80
---
> Q1. Mark (✓) against the correct answer in each of the following.  (1 x 10 = 10 Marks)
(i) What is the 56th even number?
  (a) 63  (b) 70  (c) 84  (d) 112

(ii) What is the 37th odd number?
  (a) 73  (b) 78  (c) 74  (d) 75

> Q2. Fill in the blanks.  (1 x 5 = 5 Marks)
(i) A right triangle cannot have an obtuse angle.
(ii) Sum of the acute angles of a right triangle is ______.

> SHORT ANSWER QUESTIONS
Q4. Solve the following expressions.  [2 Marks]
  (a) 175 - (58 - 28)   (b) 150 + (41 - 15)

Q14. In a triangle ABC, if 2∠A = 3∠B = 6∠C, find ∠A, ∠B and ∠C.  [3 Marks]
`;

  // Escape text for safe HTML insertion
  function esc(s) {
    return s
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  // Same auto-style token logic as docxBuilder.js, applied to escaped text.
  // Order matters: **bold** must be tried before *italic* (see docxBuilder.js).
  const INLINE_TOKEN_RE =
    /(\*\*(.+?)\*\*)|(\[[^\[\]]*\bMarks?\b[^\[\]]*\])|(\((?:[a-zA-Z]{1,5}|\d{1,3})\))|(\*(.+?)\*)/g;

  function inlineFormatHtml(text) {
    let out = "";
    let last = 0;
    let m;
    INLINE_TOKEN_RE.lastIndex = 0;
    while ((m = INLINE_TOKEN_RE.exec(text)) !== null) {
      out += esc(text.slice(last, m.index));
      if (m[1] !== undefined) {
        out += "<b>" + esc(m[2]) + "</b>";
      } else if (m[3] !== undefined) {
        out += '<span class="pv-marks">' + esc(m[3]) + "</span>";
      } else if (m[4] !== undefined) {
        out += "<b>" + esc(m[4]) + "</b>";
      } else if (m[5] !== undefined) {
        out += "<i>" + esc(m[6]) + "</i>";
      }
      last = m.index + m[0].length;
    }
    out += esc(text.slice(last));
    return out;
  }

  function renderPreview() {
    const md = editor.value;
    const lines = md.replace(/\r\n/g, "\n").split("\n");
    let html = "";

    for (const raw of lines) {
      const trimmed = raw.trim();

      if (trimmed === "") {
        html += `<div class="pv-spacer"></div>`;
        continue;
      }
      if (/^-{3,}$/.test(trimmed)) {
        html += `<hr class="pv-hr" />`;
        continue;
      }
      if (trimmed.startsWith("### ")) {
        html += `<div class="pv-classline">${esc(trimmed.slice(4))}</div>`;
        continue;
      }
      if (trimmed.startsWith("## ")) {
        html += `<div class="pv-subtitle">${esc(trimmed.slice(3))}</div>`;
        continue;
      }
      if (trimmed.startsWith("# ")) {
        html += `<div class="pv-title">${esc(trimmed.slice(2))}</div>`;
        continue;
      }
      if (trimmed.startsWith("> ")) {
        html += `<div class="pv-banner">${esc(trimmed.slice(2))}</div>`;
        continue;
      }

      const leadingSpaces = (raw.match(/^(\s*)/) || ["", ""])[1];
      const nested = leadingSpaces.length >= 2 || raw.startsWith("\t");
      html += `<div class="pv-body${nested ? " nested" : ""}">${inlineFormatHtml(trimmed)}</div>`;
    }

    preview.innerHTML = html;
  }

  function slugTitle(md) {
    const m = md.match(/^#\s+(.+)$/m);
    if (!m) return "exam";
    return m[1].trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "").slice(0, 60) || "exam";
  }

  async function handleDownload() {
    if (typeof docx === "undefined" || typeof MarkDexam === "undefined") {
      alert("The document engine hasn't finished loading yet (or failed to load — check your internet connection). Please wait a moment and try again.");
      return;
    }
    downloadBtn.disabled = true;
    const prevLabel = downloadBtn.textContent;
    downloadBtn.textContent = "Building…";
    try {
      const doc = MarkDexam.markdownToDoc(editor.value);
      const blob = await docx.Packer.toBlob(doc);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = slugTitle(editor.value) + ".docx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 2000);
    } catch (err) {
      console.error(err);
      alert("Could not build the .docx file: " + err.message);
    } finally {
      downloadBtn.disabled = false;
      downloadBtn.textContent = prevLabel;
    }
  }

  editor.addEventListener("input", renderPreview);
  sampleBtn.addEventListener("click", () => {
    editor.value = SAMPLE_MD;
    renderPreview();
  });
  downloadBtn.addEventListener("click", handleDownload);

  // initial state
  editor.value = SAMPLE_MD;
  renderPreview();
})();
