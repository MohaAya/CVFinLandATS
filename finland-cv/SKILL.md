---
name: finland-cv
description: Build or rewrite a CV for the Finnish job market as an editable Word file plus PDF - clean 1-2 page A4 layout, Responsibilities and Results under each role, ATS-safe, English or Finnish headings. Use when someone asks for a CV or resume for jobs in Finland.
---

# Finnish-style CV (Word + PDF)

Builds a CV that follows what Finnish recruiters and career services expect, and that still parses cleanly in the recruitment systems Finnish employers use (e.g. Laura, Sympa, TalentAdore, Workday). Output: `Firstname_Lastname_CV.docx` (editable) and `Firstname_Lastname_CV.pdf` (to send).

## Workflow

1. **Gather content.** If the person attached a CV, extract it (`pandoc -t plain --wrap=none cv.docx`, or `pdftotext -layout cv.pdf -`). Also pull hyperlinks from a .docx (`unzip -p cv.docx word/_rels/document.xml.rels | grep -o 'Target="http[^"]*"'`) to get the LinkedIn URL. Otherwise ask for: name, headline/field, town, phone, email, LinkedIn, profile, experience (title, employer, place, years, responsibilities, results), education, certifications, languages with levels, tools, references. Ask which language the target job ad is in: English ad -> English CV (`"lang": "en"`), Finnish ad -> Finnish CV or English CV (`"lang": "fi"` switches all headings to Finnish).
2. **Never invent facts.** Use only what the person gave. No made-up numbers, employers, dates or skills. Keep their own wording wherever it works: Finnish career services explicitly warn against AI-sounding text. Fix only typos, capitalization and clearly broken sentences. If a sentence is ambiguous, keep your best reading and tell them to check it.
3. **Split each role into Responsibilities and Results.** Responsibilities = what they owned or did. Results = what they delivered, what was adopted, recognition, numbers. If the original has no split, sort the bullets yourself and say so. Put a one-line italic `description` of the employer/project when the original has one.
4. **Apply Finnish norms:**
   - 1-2 A4 pages (1 page for early career). Reverse chronological.
   - Name alone on the first line, never "Name, PhD": parsers read the whole line as the name and fail. Put the degree in `headline` (e.g. "PhD  |  Circular Economy").
   - Contact details in the body: town + country, phone, email, LinkedIn. No date of birth, ID number, marital status, nationality or full street address.
   - No photo by default (optional in Finland, and guides disagree). Mention that the person can add one in Word if they want.
   - Languages with clear levels (native / fluent / good / basic / beginner, or CEFR A1-C2). Be honest about Finnish/Swedish level.
   - References: 1-3 people who have agreed, ideally from Finnish workplaces; otherwise `"referencesOnRequest": true`.
   - Keep layout plain: single column, no tables, text boxes, icons, skill bars or charts. The script already does this.
5. **Write `cv.json`** in the schema below, then build.
6. **Build:**
   ```bash
   # docx (npm) is usually preinstalled; if require fails: npm install docx
   node build_cv.js cv.json Firstname_Lastname_CV.docx
   soffice --headless --convert-to pdf Firstname_Lastname_CV.docx
   ```
   Use `scripts/build_cv.js` next to this file if it exists; otherwise write the script from the "Build script" section below to a file first.
7. **Verify before delivering:**
   - `pdfinfo Firstname_Lastname_CV.pdf | grep Pages` - must be 2 or fewer.
   - `pdftotext -layout Firstname_Lastname_CV.pdf - | head -40` - first line is exactly the name; sections come in order; dates sit on the same line as their job title; no garbled characters (check a, o with umlauts).
   - Render and look: `pdftoppm -png -r 70 Firstname_Lastname_CV.pdf page` then view each page for awkward breaks.
   - Over 2 pages? In this order: remove facts repeated in several places (e.g. the same award in profile, role and awards), fold education descriptions into the `org` line, merge languages and tools into one section, then tighten `margins` (minimum ~600 DXA). Never cut the person's results to save space without asking.
8. **Deliver** the .docx and .pdf. In a few lines, say what you changed and ask for missing numbers: CV checkers flag bullets without figures, so ask specific questions (team size, budget, number of clients/partners/participants, % or EUR saved, time saved) and add only what the person confirms.

## cv.json schema

```json
{
  "lang": "en",                      // "en" or "fi" (Finnish headings)
  "accent": "1F5E57",                // optional hex colour for name and headings
  "name": "First Last",
  "headline": "Degree  |  Field or target role",
  "contact": { "location": "Helsinki, Finland", "phone": "+358 ...", "email": "...", "linkedin": "https://www.linkedin.com/in/...", "website": "optional" },
  "profile": { "tagline": "optional one-line, in the person's own words", "text": "3-5 sentence profile" },
  "competencies": ["Skill one", "Skill two"],
  "experience": [
    { "title": "Job title", "dates": "2022 – present", "org": "Employer, City",
      "description": "optional one line about the employer or project",
      "responsibilities": ["..."], "results": ["..."], "bullets": ["optional, used when no split"] }
  ],
  "extraSections": [ { "heading": "Teaching and Training", "bullets": ["..."], "lines": ["optional plain lines"] } ],
  "education": [ { "degree": "M.Sc., Field", "year": "2019", "org": "University, Country", "description": "optional" } ],
  "certifications": ["Certificate – Issuer, year"],
  "languages": ["Finnish – native", "English – fluent"],
  "skills": [ { "label": "Tools", "text": "..." }, { "label": "Publications", "text": "..." } ],
  "references": ["Name, Organization  |  email"],
  "referencesOnRequest": false,
  "margins": { "top": 750, "bottom": 650, "left": 1000, "right": 1000 }
}
```
(Comments above are for explanation only; real JSON must not contain them.) Omit any section the person has no content for.

## Build script

Save as `build_cv.js` if `scripts/build_cv.js` is not available:

```javascript
// Usage: node build_cv.js cv.json [output.docx]
const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, ExternalHyperlink, AlignmentType, LevelFormat,
  BorderStyle, Tab, TabStopType } = require('docx');

const cv = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const out = process.argv[3] || 'CV.docx';
const M = cv.margins || { top: 750, bottom: 650, left: 1000, right: 1000 }; // DXA, A4 page is 11906 wide
const ACCENT = (cv.accent || '1F5E57').replace('#', ''), GREY = '555555', B = 20; // 10pt body
const L = cv.lang === 'fi'
  ? { profile: 'Profiili', comp: 'Osaaminen', exp: 'Työkokemus', resp: 'Vastuualueet', res: 'Tulokset', edu: 'Koulutus',
      cert: 'Sertifikaatit', lang: 'Kielitaito', skills: 'Kielitaito ja muu osaaminen', refs: 'Suosittelijat', onreq: 'Suosittelijat pyydettäessä.' }
  : { profile: 'Profile', comp: 'Core Competencies', exp: 'Work Experience', resp: 'Responsibilities', res: 'Results', edu: 'Education',
      cert: 'Certifications', lang: 'Languages', skills: 'Languages and Skills', refs: 'References', onreq: 'References available on request.' };

const heading = (t) => new Paragraph({ keepNext: true, spacing: { before: 130, after: 50 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 2 } },
  children: [new TextRun({ text: t.toUpperCase(), bold: true, color: ACCENT, size: 23, characterSpacing: 10 })] });
const para = (t, o = {}) => new Paragraph({ spacing: { after: 40, line: 250 }, children: [new TextRun({ text: t, size: B, ...o })] });
const bullet = (t) => new Paragraph({ numbering: { reference: 'b', level: 0 }, spacing: { after: 10, line: 240 }, children: [new TextRun({ text: t, size: B })] });
const sub = (t) => new Paragraph({ keepNext: true, spacing: { before: 30, after: 10 }, children: [new TextRun({ text: t, bold: true, size: B, color: ACCENT })] });
const labelLine = (k, v) => new Paragraph({ spacing: { after: 50, line: 250 }, children: [new TextRun({ text: k + ': ', bold: true, size: B }), new TextRun({ text: v, size: B })] });
const entry = (title, dates, org, desc) => {
  const r = [
    new Paragraph({ keepNext: true, tabStops: [{ type: TabStopType.RIGHT, position: 11906 - M.left - M.right }], spacing: { before: 80, after: 0 }, children: [
      new TextRun({ text: title, bold: true, size: 22 }), new TextRun({ children: [new Tab()] }), new TextRun({ text: dates || '', bold: true, size: 22 })] }),
    new Paragraph({ keepNext: true, spacing: { after: desc ? 20 : 40 }, children: [new TextRun({ text: org || '', color: GREY, size: B })] }),
  ];
  if (desc) r.push(new Paragraph({ keepNext: true, spacing: { after: 40 }, children: [new TextRun({ text: desc, italics: true, size: B })] }));
  return r;
};
const link = (text, url) => new ExternalHyperlink({ link: url, children: [new TextRun({ text, size: B, color: ACCENT })] });
const strip = (u) => u.replace(/^https?:\/\/(www\.)?/, '').replace(/\/$/, '');

const c = cv.contact || {}, sep = () => new TextRun({ text: '  |  ', size: B });
const contact = [];
[c.location, c.phone].filter(Boolean).forEach(t => { if (contact.length) contact.push(sep()); contact.push(new TextRun({ text: t, size: B })); });
if (c.email) { if (contact.length) contact.push(sep()); contact.push(link(c.email, 'mailto:' + c.email)); }
[c.linkedin, c.website].filter(Boolean).forEach(u => { if (contact.length) contact.push(sep()); contact.push(link(strip(u), u)); });

const ch = [
  new Paragraph({ spacing: { after: 20 }, children: [new TextRun({ text: cv.name, bold: true, size: 40, color: ACCENT })] }),
];
if (cv.headline) ch.push(new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: cv.headline, size: 24, bold: true, color: GREY })] }));
ch.push(new Paragraph({ spacing: { after: 60 }, children: contact }));

if (cv.profile) {
  ch.push(heading(L.profile));
  if (cv.profile.tagline) ch.push(para(cv.profile.tagline, { italics: true }));
  if (cv.profile.text) ch.push(para(cv.profile.text));
}
if (cv.competencies?.length) { ch.push(heading(L.comp)); ch.push(para(cv.competencies.join(' · '))); }
if (cv.experience?.length) {
  ch.push(heading(L.exp));
  for (const e of cv.experience) {
    ch.push(...entry(e.title, e.dates, e.org, e.description));
    if (e.responsibilities?.length) { ch.push(sub(L.resp)); e.responsibilities.forEach(t => ch.push(bullet(t))); }
    if (e.results?.length) { ch.push(sub(L.res)); e.results.forEach(t => ch.push(bullet(t))); }
    (e.bullets || []).forEach(t => ch.push(bullet(t)));
  }
}
for (const s of cv.extraSections || []) { ch.push(heading(s.heading)); (s.bullets || []).forEach(t => ch.push(bullet(t))); (s.lines || []).forEach(t => ch.push(para(t))); }
if (cv.education?.length) { ch.push(heading(L.edu)); cv.education.forEach(e => ch.push(...entry(e.degree, e.year, e.org, e.description))); }
if (cv.certifications?.length) { ch.push(heading(L.cert)); cv.certifications.forEach(t => ch.push(para(t))); }
if (cv.languages?.length || cv.skills?.length) {
  ch.push(heading(cv.skills?.length ? L.skills : L.lang));
  if (cv.languages?.length) ch.push(labelLine(L.lang, cv.languages.join(', ')));
  (cv.skills || []).forEach(s => ch.push(labelLine(s.label, s.text)));
}
if (cv.references?.length) { ch.push(heading(L.refs)); cv.references.forEach(t => ch.push(para(t))); }
else if (cv.referencesOnRequest) { ch.push(heading(L.refs)); ch.push(para(L.onreq)); }

const doc = new Document({
  creator: cv.name, title: cv.name + ' CV',
  styles: { default: { document: { run: { font: 'Calibri', size: B } } } },
  numbering: { config: [{ reference: 'b', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 300, hanging: 220 } } } }] }] },
  sections: [{ properties: { page: { margin: M } }, children: ch }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(out, b); console.log('wrote ' + out); });
```
