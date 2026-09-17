// ============================================================
// MASTER COVER LETTER TEMPLATE
// Career Accelerator Project
// ============================================================
// HOW TO USE:
// 1. Replace all CANDIDATE.* and LETTER.* variables below
// 2. Write the five paragraphs in the CONTENT block
// 3. Run: node cover_letter_template.js
// 4. Output: candidate_cover_letter.docx
// ============================================================

const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  BorderStyle,
  ExternalHyperlink,
} = require("docx");
const fs = require("fs");

// ============================================================
// CONFIG — replace all values for each new candidate
// ============================================================

const CANDIDATE = {
  name: "Full Name",
  email: "email@example.com",
  phone: "000-000-0000",
  location: "City, Country",
  linkedin: "https://linkedin.com/in/username",
  outputFile: "candidate_cover_letter.docx",
};

const LETTER = {
  date: "20 March 2026",
  addressee: "Hiring Team",
  company: "Company Name",
  location: "Remote",
  subject: "Application: [Role Title] — [Department or Project Name]",

  // Five paragraphs — plain text, no special characters
  // No em dashes, no smart quotes, no ellipsis, no arrows
  // Each paragraph should be 3 to 5 sentences maximum

  para1:
    "Who you are and why you are applying. Mention the exact role title. Mention contractor availability if relevant. State you are available immediately.",

  para2:
    "Your proof paragraph. Name your two strongest and most relevant experiences or projects. Frame them as live shipped work, not tutorials. Include one specific technical detail from each.",

  para3:
    "Your unique differentiator. What sets you apart that most applicants will not have. Be specific. This should be the paragraph the hiring manager remembers.",

  para4:
    "Your stack match. Confirm the key technologies from the JD that you use. Mention one skill you are actively developing and be honest about it.",

  para5:
    "Short confident close. Reference the specific next step in their hiring process if known (AI interview, technical test etc). One or two sentences only.",
};

// ============================================================
// STYLES — do not change
// ============================================================

const FONT = "Times New Roman";
const DARK = "1a1a1a";
const GRAY = "444444";
const LGRAY = "666666";
const BLUE = "1a5276";

// ============================================================
// HELPERS — do not change
// ============================================================

const r = (text, o = {}) =>
  new TextRun({
    text,
    font: FONT,
    size: o.s || 21,
    color: o.c || DARK,
    bold: !!o.b,
    italics: !!o.i,
    underline: o.u || undefined,
  });

const line = (ch, before = 0, after = 180) =>
  new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before, after },
    children: Array.isArray(ch) ? ch : [r(ch)],
  });

const bodyPara = (text, before = 200) =>
  new Paragraph({
    spacing: { before, after: 0 },
    children: [r(text, { s: 20, c: GRAY })],
  });

// ============================================================
// BUILD DOCUMENT
// ============================================================

const doc = new Document({
  sections: [
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1080, right: 1260, bottom: 1080, left: 1260 },
        },
      },
      children: [
        // NAME
        new Paragraph({
          spacing: { before: 0, after: 56 },
          children: [r(CANDIDATE.name, { s: 48 })],
        }),

        // CONTACT
        new Paragraph({
          spacing: { before: 0, after: 60 },
          children: [
            r(
              `${CANDIDATE.email}  |  ${CANDIDATE.phone}  |  ${CANDIDATE.location}  |  `,
              { s: 19, c: LGRAY },
            ),
            new ExternalHyperlink({
              link: CANDIDATE.linkedin,
              children: [r("LinkedIn", { s: 19, c: BLUE, u: {} })],
            }),
          ],
        }),

        // RULE
        new Paragraph({
          spacing: { before: 60, after: 200 },
          border: {
            bottom: {
              style: BorderStyle.SINGLE,
              size: 4,
              color: "000000",
              space: 1,
            },
          },
          children: [],
        }),

        // DATE
        line(LETTER.date, 0, 160),

        // ADDRESSEE
        line(
          [r(`${LETTER.addressee}, `, {}), r(LETTER.company, { b: true })],
          0,
          40,
        ),
        line(LETTER.location, 0, 200),

        // SUBJECT
        new Paragraph({
          spacing: { before: 0, after: 200 },
          children: [r(LETTER.subject, { b: true })],
        }),

        // SALUTATION
        line("Dear Hiring Team,", 0, 200),

        // FIVE BODY PARAGRAPHS
        bodyPara(LETTER.para1),
        bodyPara(LETTER.para2, 200),
        bodyPara(LETTER.para3, 200),
        bodyPara(LETTER.para4, 200),
        bodyPara(LETTER.para5, 200),

        // SIGN OFF
        line("Yours sincerely,", 280, 280),

        // SIGNATURE NAME
        new Paragraph({
          spacing: { before: 0, after: 56 },
          children: [r(CANDIDATE.name, { s: 21 })],
        }),

        // SIGNATURE CONTACT
        new Paragraph({
          spacing: { before: 0, after: 0 },
          children: [
            r(`${CANDIDATE.email}  |  ${CANDIDATE.phone}  |  `, {
              s: 19,
              c: LGRAY,
            }),
            new ExternalHyperlink({
              link: CANDIDATE.linkedin,
              children: [r("LinkedIn", { s: 19, c: BLUE, u: {} })],
            }),
          ],
        }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(CANDIDATE.outputFile, buf);
  console.log(`Cover letter saved: ${CANDIDATE.outputFile}`);
});
