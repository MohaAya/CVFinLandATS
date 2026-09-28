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
