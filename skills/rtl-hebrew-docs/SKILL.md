---
name: rtl-hebrew-docs
description: "When the user wants a Hebrew (RTL) document as PDF, Word, or slides: price proposals (הצעת מחיר), contracts, meeting minutes, formal letters, certificates, reports, or asks why Hebrew shows as boxes or runs left-to-right. Default to HTML with dir=\"rtl\" printed to PDF by headless Chrome; use DOCX only for editing in Word. Never issue tax invoices or receipts (חשבונית מס, קבלה): those need approved software. For proposal content, see sales-enablement. For the copy itself, see copywriting."
license: MIT
compatibility: Headless Chrome or Chromium for the primary PDF path, plus a Hebrew-capable system font (macOS and Windows ship one; on Linux check fc-list :lang=he). Python 3.9+ with python-docx for editable Word files. Node.js with pptxgenjs for slides. reportlab and python-bidi only for the batch PDF fallback.
metadata:
  version: 1.0.0
---

# Hebrew Document Generator

Generate clean, professional **right-to-left Hebrew documents** that render correctly **everywhere**, not just in the one viewer you happened to test.

> **The single most important rule, learned the hard way:** A Hebrew document can have perfect RTL markup and *still* display left-to-right in Pages, Quick Look, Google Drive preview, or LibreOffice. The only path that renders identically across every viewer is **HTML (`dir="rtl"`) → headless Chrome → PDF**. Make that your default. Reach for DOCX only when the recipient must edit the file in Microsoft Word.

---

## Step 1: Pick the output by what the recipient actually needs

| The recipient needs… | Use | Why this one |
|---|---|---|
| A finished document that looks identical everywhere **(DEFAULT)** | **HTML `dir="rtl"` → headless Chrome → PDF** | Native bidi, real CSS, and the fonts are embedded in the PDF. This is the only path that never silently breaks RTL. |
| An **editable** file they will open **in Microsoft Word** | `python-docx` with full RTL markers | The *only* reason to choose DOCX. ⚠ Its RTL is not honored by Pages / Quick Look / Google Drive preview / LibreOffice, verify in real Word. |
| A slide deck / presentation | `pptxgenjs` (Node) with `rtlMode` | Native RTL text boxes and bullets. |
| Hundreds of programmatic PDFs with no HTML/Chrome available | `reportlab` + `python-bidi` (`base_dir='R'`) | Low-level fallback only. Manual layout, easy to get wrong, avoid unless Chrome is unavailable. |

**Intentionally demoted: WeasyPrint.** It needs Cairo/Pango system libraries and fails to install cleanly on many macOS setups. If a teammate already has it working, its CSS support is fine, but do not make it the default and do not assume it is installed.

> Full engine reference, bidi rules, and font traps: **`references/rtl-rendering.md`**.

---

## Step 2: Install only what the chosen path needs

```bash
# Primary path (HTML → PDF): nothing to pip-install. Just confirm Chrome/Chromium exists.
#   macOS:   "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --version
#   Linux:   google-chrome --version  ||  chromium --version
#   Windows: where chrome

# Editable Word files:
pip install python-docx

# Slides:
npm install pptxgenjs

# Low-level batch PDF fallback only:
pip install reportlab python-bidi
```

Chrome uses the operating system's fonts. macOS and Windows ship Hebrew fonts, so the primary path works out of the box. On **Linux**, check with `fc-list :lang=he`; if nothing good is listed, install Noto (`sudo apt-get install fonts-noto-core` on Debian/Ubuntu, `google-noto-sans-hebrew-fonts` on Fedora) or copy an OFL TTF into `~/.fonts` and run `fc-cache -fv`. If Chrome is installed somewhere unusual, set `CHROME_PATH` to its binary.

---

## Step 3: Generate a Hebrew PDF (the default path)

Write semantic HTML with `dir="rtl"`, then render it with the bundled helper. The helper finds Chrome automatically, renders to PDF, and strips the browser header/footer.

```bash
python3 scripts/html_to_pdf.py --html document.html --output document.pdf
```

Minimal RTL-correct HTML, note `dir="rtl"`, a Hebrew-capable font stack, generous line-height, and **logical** CSS properties (`margin-inline`, `text-align: start`) so the layout mirrors automatically:

```html
<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
  <meta charset="utf-8">
  <style>
    @page { size: A4; margin: 20mm; }
    body {
      font-family: "Rubik", "Heebo", "Arial Hebrew", "Noto Sans Hebrew", sans-serif;
      direction: rtl;
      font-size: 12pt;
      line-height: 1.7;          /* Hebrew needs more leading than Latin */
      color: #1a1a2e;
    }
    h1 { font-size: 22pt; margin-block-end: 0.4em; }
    .meta { color: #555; }
    table { width: 100%; border-collapse: collapse; margin-block: 1em; }
    th, td { border: 1px solid #ccc; padding: 8px 12px; text-align: start; }
    /* Keep prices/numbers LTR inside an RTL row so they never reorder */
    .num { direction: ltr; unicode-bidi: isolate; text-align: start; }
  </style>
</head>
<body>
  <h1>הצעת מחיר</h1>
  <p class="meta">סטודיו רקפת · עיצוב גרפי · 30 ביוני 2026</p>
  <p>…</p>
</body>
</html>
```

The same helper can build a couple of ready templates from JSON instead of hand-written HTML:

```bash
python3 scripts/html_to_pdf.py --template proposal --data proposal.json --output proposal.pdf
python3 scripts/html_to_pdf.py --template certificate --data cert.json --output certificate.pdf
```

**Why HTML wins for Hebrew:** the browser's bidi engine handles mixed Hebrew/English/number runs correctly, tables mirror automatically, `@font-face` lets you embed a brand font, and the PDF carries its fonts, so it looks the same on every machine that opens it.

---

## Step 4: Generate an editable Hebrew Word file (only when they need Word)

Use this **only** when the recipient will edit the document in Microsoft Word. For RTL to survive, you must set bidi at three levels: section, every paragraph, and every run (with `w:rtl` **and** `w:cs` for the complex-script slot). Word formats Hebrew from the complex-script properties, so size and bold must also go on `w:szCs` and `w:bCs`, or Hebrew stays at the default size and never turns bold.

```bash
python3 scripts/make_docx.py --type contract --output contract.docx
```

```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt

def set_rtl_paragraph(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    pPr.insert(0, OxmlElement('w:bidi'))         # paragraph direction = RTL; start = right margin

def set_rtl_run(run, size_pt, bold=False):
    run.font.size = Pt(size_pt); run.bold = bold
    run.font.cs_bold = bold                      # w:bCs: bold for Hebrew
    rPr = run._r.get_or_add_rPr()
    szCs = OxmlElement('w:szCs'); szCs.set(qn('w:val'), str(int(size_pt * 2)))
    rPr.find(qn('w:sz')).addnext(szCs)           # w:szCs: size for Hebrew
    rPr.append(OxmlElement('w:rtl'))             # run is RTL (append last: schema order)
    rPr.append(OxmlElement('w:cs'))              # apply to complex-script font slot
```

⚠ **Reality check before you ship a DOCX:** open it in the recipient's actual tool. python-docx RTL is reliably honored only by **Microsoft Word**. Pages, macOS Quick Look, Google Drive preview, and LibreOffice frequently ignore the markers and show the text LTR. If the recipient uses Google Docs, upload the DOCX (or an HTML file) and let Google convert it. If they just need *a document*, give them the PDF from Step 3 instead.

> Full python-docx RTL recipe (headings, tables, page setup): **`scripts/make_docx.py`**.

---

## Step 5: Generate a Hebrew presentation

```bash
node scripts/make_pptx.js
```

```javascript
const pptxgen = require("pptxgenjs");
const pptx = new pptxgen();
pptx.layout = "LAYOUT_16x9";
pptx.rtlMode = true;                              // deck-wide RTL

const slide = pptx.addSlide();
slide.addText("פתיחת הכנס השנתי", {
  x: 0.5, y: 0.5, w: "90%", h: 1.0,
  fontSize: 28, fontFace: "Rubik", bold: true,
  align: "right", rtlMode: true,
});
slide.addText(
  [
    { text: "מגמות בשוק לשנה הקרובה", options: { bullet: true, rtlMode: true } },
    { text: "הצגת המוצרים החדשים", options: { bullet: true, rtlMode: true } },
    { text: "מפגש עם השותפים", options: { bullet: true, rtlMode: true } },
  ],
  { x: 0.5, y: 2.0, w: "90%", h: 3.0, fontSize: 18, fontFace: "Rubik", align: "right", rtlMode: true }
);
pptx.writeFile({ fileName: "conference-opening.pptx" });
```

Set `rtlMode: true` at **both** the deck level and every text block, pptxgenjs does not inherit it. Use a Hebrew-capable `fontFace` (Rubik / Heebo / Arial); the renderer falls back to boxes otherwise.

---

## Step 6: VERIFY the output (do not skip this)

Generation succeeding is **not** proof the Hebrew renders. The classic failure is a file that opens with Hebrew flowing left-to-right or showing as boxes. Always confirm:

1. **Open it.** For a PDF, render page 1 to an image or extract its text and read it back:
   ```bash
   # Render page 1 to an image and look at it
   pdftoppm -r 80 -png -f 1 -l 1 document.pdf page      # Linux/macOS (poppler)
   qlmanage -t -s 1200 -o /tmp document.pdf               # macOS thumbnail
   # Or extract text to confirm real Hebrew (not '?????' or boxes):
   python3 -c "import pdfplumber,sys; print(pdfplumber.open(sys.argv[1]).pages[0].extract_text())" document.pdf
   ```
   Text extraction often returns Hebrew in visual order (words reversed). That is an extraction artifact, not a broken PDF: judge direction from the image.
2. **Confirm three things:** (a) Hebrew letters render as letters, not □ boxes or `?`; (b) text flows right-to-left; (c) any embedded numbers, prices, dates, or English words sit in the right place and didn't reorder.
3. **DOCX only:** open in the recipient's real application (ideally Word), not just Quick Look.

If Hebrew shows as boxes → the font lacks Hebrew glyphs (see Gotchas). If it flows LTR → the direction/bidi setting didn't take.

---

## Safe document templates (this skill's scope)

These are **commercial / administrative** documents, not regulated accounting records, so they are safe to generate freely. Field specs and standard Hebrew phrasing: **`references/document-templates.md`**.

| Template | Hebrew | Typical fields |
|---|---|---|
| Price proposal | הצעת מחיר | Business details, recipient, itemized lines, subtotal, VAT, total, validity period, terms |
| Service contract | חוזה / הסכם התקשרות | Parties + IDs, scope, payment terms, duration, termination, confidentiality, signatures |
| Meeting minutes | פרוטוקול ישיבה | Date/time, attendees, agenda, discussion, numbered decisions, action items + owners |
| Formal letter | מכתב רשמי | Sender letterhead, date, addressee, subject, body, sign-off |
| Certificate | תעודה / אישור | Recipient name, achievement/confirmation text, issuer, date, signature line |
| Report | דוח | Title, period, sections, tables/figures, summary |

> **⚠ Out of scope, tax invoices and receipts.** In Israel a `חשבונית מס` (tax invoice) and `קבלה` (receipt) are regulated accounting documents. They must be issued from Tax-Authority-approved bookkeeping software that provides mandatory sequential numbering, the `קובץ אחיד` (BKMVDATA) export, and, above the statutory threshold, an **allocation number (מספר הקצאה)** obtained live from the Tax Authority under the *Israel Invoice* reform. A PDF produced by a script has none of these, so it is **not a valid tax invoice** and the buyer cannot deduct input VAT from it. Do not generate these here, route them through approved invoicing software. (A *price proposal* showing a VAT line is fine; it is not an accounting document.)
>
> **VAT note:** if a proposal shows VAT, treat the rate as a **parameter** (currently **18%** in Israel, raised from 17% on 1 Jan 2025). Never hardcode it, it changes by law.

---

## Examples

### Example 1: Price proposal as PDF
**User:** "תכין הצעת מחיר בעברית לסטודיו עיצוב, כ-PDF."
**Do:** Build RTL HTML with the studio's header, an itemized table (description / qty / unit price / line total), subtotal, a VAT line (18%, parameterized), total, a 30-day validity note, and a Hebrew-capable font. Render with `scripts/html_to_pdf.py`. Then **verify** the PDF opens RTL with real Hebrew.

### Example 2: Service contract as editable Word
**User:** "אני צריך חוזה התקשרות בעברית שאוכל לערוך ב-Word."
**Do:** Use `scripts/make_docx.py --type contract`. Set section + paragraph + run bidi, David or Heebo font, structured sections (צדדים, היקף, תנאי תשלום, תקופה, ביטול, סודיות, חתימות). Tell the user explicitly to open it in **Microsoft Word**, and that if they only need to read/print it, the PDF path is more reliable.

### Example 3: Committee meeting minutes as PDF
**User:** "תפיק פרוטוקול של ישיבת ועד הבית."
**Do:** RTL HTML with header (date, time, location, attendees), an agenda list, a discussion section, **numbered decisions (החלטות)**, and an action-items table (משימה / אחראי / תאריך יעד). Render to PDF and verify.

### Example 4: Conference presentation
**User:** "תכין מצגת פתיחה לכנס בעברית."
**Do:** `scripts/make_pptx.js` with deck-level `rtlMode`, a Rubik title slide, right-aligned RTL bullet slides, and a clean color scheme. Open the file to confirm bullets are right-aligned RTL.

### Example 5: Participation certificate as PDF
**User:** "תעצב תעודת השתתפות בקורס בעברית."
**Do:** A centered, bordered RTL HTML certificate, recipient name large, confirmation sentence, issuer + date + signature line, rendered to PDF. Great showcase of centered RTL typography; verify the name and date render correctly.

---

## Gotchas (every one of these has bitten a real document)

- **A font without Hebrew glyphs renders as boxes (□).** Roboto, many decorative Latin fonts, and the reportlab default (Helvetica) have **no** Hebrew. Use Rubik, Heebo, Assistant, Frank Ruhl Libre, David, or Arial Hebrew. Verify the font covers the Hebrew Unicode block `U+0590-U+05FF`.
- **`get_display()` without `base_dir='R'` breaks lines that start in English.** A line like `"OpenAI הוא כלי"` is mis-ordered when bidi auto-detects LTR from the first character. In reportlab/python-bidi always pass `get_display(text, base_dir='R')`. (HTML/Chrome and Word do this correctly on their own, another reason to prefer them.)
- **Never put arrow glyphs (→ ← ↑ ↓ « ») in RTL body text.** Mail clients and some PDF viewers re-orient them unpredictably. Use commas, "ואז", colons, or numbered lists instead.
- **Numbers, prices, dates, and English fragments inside Hebrew need isolation.** In HTML wrap them in a span with `unicode-bidi: isolate; direction: ltr;`. In reportlab they're handled by `base_dir='R'`. Otherwise `2026 ₪1,200` can land in the wrong order.
- **Do not add nikud (vowel marks) to business/legal Hebrew.** It is not used in formal documents and reads as unprofessional. Agents sometimes add it "for clarity", don't.
- **Dates are DD/MM/YYYY** in Israeli secular documents (agents default to MM/DD/YYYY). Use Hebrew-calendar dates only for explicitly religious/traditional documents.
- **Sending Hebrew through `curl -d` on Windows/Git-Bash mangles it to `?????`.** If a script posts Hebrew to an API, write the JSON to a UTF-8 file and send with `curl --data-binary @file.json`, then confirm the stored text is real Hebrew.

---

## Troubleshooting

**Hebrew shows as boxes or `?`** → The font has no Hebrew glyphs or wasn't found. HTML path: pick a Hebrew font in the stack and (Linux) install `fonts-noto-core` or another Hebrew font, then confirm with `fc-list :lang=he`. reportlab: register a Hebrew TTF with `pdfmetrics.registerFont(TTFont(...))`, never rely on the Helvetica default.

**Text flows left-to-right** → The direction setting didn't apply. HTML: confirm `dir="rtl"` on `<html>` and `direction: rtl` in CSS. DOCX: confirm section `w:bidi`, paragraph `w:bidi`, and run `w:rtl`+`w:cs` are all present (and `w:szCs`/`w:bCs` if size or bold look wrong), and open in **Word**, not a previewer. reportlab: you forgot `get_display(text, base_dir='R')`.

**Numbers/punctuation land in the wrong spot** → Bidi isolation missing. HTML: `unicode-bidi: isolate` on LTR spans. reportlab: ensure `base_dir='R'`.

**DOCX looks right in Word but wrong in Google Drive / Pages** → Expected. Those previewers ignore python-docx RTL. Deliver a PDF, or upload to Google Docs and let it convert.

**WeasyPrint won't install / crashes on import** → It needs Cairo/Pango. Don't fight it, switch to the headless-Chrome path, which needs no extra libraries.

---

## Bundled resources

### Scripts
- `scripts/html_to_pdf.py`, **Primary generator.** Renders an HTML file (or a built-in `proposal` / `certificate` template fed from JSON) to PDF via headless Chrome, auto-detecting the Chrome/Chromium binary across macOS/Linux/Windows (or `CHROME_PATH`) and stripping the browser header/footer. Falls back to WeasyPrint only if explicitly requested and installed. `python3 scripts/html_to_pdf.py --help`
- `scripts/make_docx.py`, Editable Hebrew Word generator with correct three-level RTL (section + paragraph + run), complex-script size and bold, A4 page, and Hebrew fonts; builds a sample `contract` or `letter`. Carries the in-code caveat about non-Word previewers. `python3 scripts/make_docx.py --help`
- `scripts/make_pptx.js`, Hebrew presentation generator with deck-wide and per-block `rtlMode` and right-aligned bullets. `node scripts/make_pptx.js`

### References
- `references/rtl-rendering.md`, The engineering core: engine comparison and ranking, the bidi rules, mixed Hebrew/English/number handling, the headless-Chrome command, and font-glyph traps.
- `references/hebrew-fonts.md`, Hebrew font catalog by document type, which fonts ship on which OS, redistribution-safe (OFL) fonts you may bundle, pairing suggestions, and typography settings.
- `references/document-templates.md`, Field specifications and standard Hebrew phrasing for the safe document types (proposal, contract, minutes, letter, certificate, report), plus the explicit note that tax invoices/receipts are out of scope and why.

---

## Related Skills

- **sales-enablement**: What goes into a proposal, one-pager, or deck before it is laid out in Hebrew
- **copywriting**: Writing the Hebrew page or proposal copy itself
- **pricing**: Deciding the prices and packages a proposal quotes
