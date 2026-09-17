// ============================================================
// MASTER CV TEMPLATE
// Career Accelerator Project
// ============================================================
// HOW TO USE:
// 1. Replace all CANDIDATE.* variables in the CONFIG block below
// 2. Replace EXPERIENCES, PROJECTS, EDUCATION arrays with real data
// 3. Run: node cv_template.js
// 4. Output: candidate_cv.docx
// ============================================================
 
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  LevelFormat,
  BorderStyle,
  ExternalHyperlink,
  TabStopType,
} = require("docx");
const fs = require("fs");
 
// ============================================================
// CONFIG — replace all values below for each new candidate
// ============================================================
 
const CANDIDATE = {
  name: "Full Name",
  email: "email@example.com",
  phone: "000-000-0000",
  location: "City, Country",
  linkedin: "https://linkedin.com/in/username",
  portfolio: "https://portfolio.vercel.app",
  roleTitle: "Software Engineer (Front-end)", // must match job advert title exactly
  stackLine: "React  |  Next.js  |  TypeScript  |  Vanilla JS", // key stack keywords
  summary:
    "Write 3 to 4 sentence summary here. Mention role title, key stack, most impressive project or employer, unique differentiator, and contractor availability if relevant.",
  outputFile: "candidate_cv.docx",
};
 
// Skills — each row is { label, value }
const SKILLS = [
  {
    label: "Languages & Web",
    value: "JavaScript ES6+, TypeScript, HTML5, CSS3, SCSS, BEM",
  },
  { label: "Frameworks & Libraries", value: "React, Next.js, Redux, Axios" },
  {
    label: "Frontend Specialisms",
    value:
      "Client-side web apps, state management, real-time data, UI interactivity, performance optimization, debugging",
  },
  {
    label: "Tools & Platforms",
    value: "Firebase v9, Vercel, Git, GitHub, REST APIs, npm, Yarn",
  },
  {
    label: "AI / Data Skills",
    value:
      "Technical prompt writing, binary rubric creation, screenshot-based UI verification",
  }, // remove if not relevant
  { label: "In Development", value: "Canvas API, Jest, React Testing Library" }, // honest skills in progress — remove if none
];
 
// Experience — most recent first
// Each role: { title, company, location, dates, bullets: [] }
const EXPERIENCES = [
  {
    title: "Frontend Developer Intern",
    company: "Company Name",
    location: "Remote",
    dates: "Jan 2026 – Present",
    bullets: [
      "Bullet point one — lead with what you built and deployed.",
      "Bullet point two — API integration, real-time data, async state.",
      "Bullet point three — performance, debugging, production stability.",
    ],
  },
  {
    title: "Founder & Software Engineer (Front-end)",
    company: "Project or Company Name",
    location: "",
    dates: "Jan 2024 – Present",
    bullets: [
      "Bullet point one — what you architected and shipped.",
      "Bullet point two — technical feature with measurable outcome.",
      "Bullet point three — performance or engagement metric.",
      "Bullet point four — responsive design, accessibility, or testing.",
    ],
  },
  {
    title: "Other Role",
    company: "Other Company",
    location: "",
    dates: "Oct 2023 – Present",
    bullets: [
      "One bullet only for non-technical roles — connect it to a skill the job advert needs.",
    ],
  },
];
 
// Projects — most relevant first, max 3 to 4
const PROJECTS = [
  {
    title: "Project Name",
    demo: "https://livesite.com",
    github: "https://github.com/username/repo",
    bullet:
      "One line description: what it is, tech stack, key technical feature.",
  },
  {
    title: "Project Name 2",
    demo: "https://livesite2.com",
    github: "https://github.com/username/repo2",
    bullet:
      "One line description: what it is, tech stack, key technical feature.",
  },
  {
    title: "Project Name 3",
    demo: "https://livesite3.com",
    github: "https://github.com/username/repo3",
    bullet:
      "One line description: what it is, tech stack, key technical feature.",
  },
];
 
// Education — most recent first
const EDUCATION = [
  {
    institution: "Institution Name",
    qualification: "Degree or Certificate Title",
    dates: "2020 – 2024",
  },
];
 
// ============================================================
// STYLES — do not change these unless redesigning
// ============================================================
 
const FONT = "Times New Roman";
const DARK = "1a1a1a";
const GRAY = "444444";
const LGRAY = "777777";
const BLUE = "1a5276";
 
// ============================================================
// HELPERS — do not change
// ============================================================
 
const r = (text, o = {}) =>
  new TextRun({
    text,
    font: FONT,
    size: o.s || 19,
    color: o.c || DARK,
    bold: !!o.b,
    italics: !!o.i,
    underline: o.u || undefined,
  });
 
const sep = () => r("  |  ", { s: 19, c: LGRAY });
 
function heading(text) {
  return new Paragraph({
    spacing: { before: 160, after: 70 },
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000", space: 1 },
    },
    children: [r(text, { b: true, s: 21, c: DARK })],
  });
}
 
function roleHeader(title, company, dates) {
  return new Paragraph({
    spacing: { before: 120, after: 20 },
    tabStops: [{ type: TabStopType.RIGHT, position: 9360 }],
    children: [
      r(company, { b: true, s: 20 }),
      r(`\t${dates}`, { s: 19, c: LGRAY }),
    ],
  });
}
 
function roleSubtitle(text) {
  return new Paragraph({
    spacing: { before: 0, after: 26 },
    children: [r(text, { i: true, s: 18, c: GRAY })],
  });
}
 
function bullet(text, ref) {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { before: 28, after: 28 },
    children: [r(text, { s: 18, c: DARK })],
  });
}
 
function skillRow(label, value) {
  return new Paragraph({
    spacing: { before: 32, after: 28 },
    children: [
      r(label + ": ", { b: true, s: 19 }),
      r(value, { s: 19, c: GRAY }),
    ],
  });
}
 
function projHeader(title, demoUrl, githubUrl) {
  return new Paragraph({
    spacing: { before: 110, after: 22 },
    children: [
      r(title + "  ", { b: true, s: 19 }),
      r("| ", { s: 18, c: LGRAY }),
      new ExternalHyperlink({
        link: demoUrl,
        children: [r("Live Demo", { s: 18, c: BLUE, u: {} })],
      }),
      r("  |  ", { s: 18, c: LGRAY }),
      new ExternalHyperlink({
        link: githubUrl,
        children: [r("GitHub", { s: 18, c: BLUE, u: {} })],
      }),
    ],
  });
}
 
// ============================================================
// BUILD BULLET REFERENCES — one per role/section
// ============================================================
 
const totalBulletRefs = EXPERIENCES.length + PROJECTS.length + 2; // +2 for skills and education buffers
const bRefs = Array.from({ length: totalBulletRefs }, (_, i) => ({
  reference: `b${i}`,
  levels: [
    {
      level: 0,
      format: LevelFormat.BULLET,
      text: "\u2022",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 360, hanging: 200 } } },
    },
  ],
}));
 
// ============================================================
// BUILD DOCUMENT
// ============================================================
 
const children = [];
 
// NAME
children.push(
  new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 0, after: 50 },
    children: [r(CANDIDATE.name, { s: 52 })],
  }),
);
 
// SUBTITLE
children.push(
  new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 0, after: 60 },
    children: [
      r(`${CANDIDATE.roleTitle}  |  ${CANDIDATE.stackLine}`, {
        s: 20,
        c: GRAY,
      }),
    ],
  }),
);
 
// CONTACT
children.push(
  new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 0, after: 60 },
    children: [
      r(
        `${CANDIDATE.email}  |  ${CANDIDATE.phone}  |  ${CANDIDATE.location}  |  `,
        { s: 18, c: GRAY },
      ),
      new ExternalHyperlink({
        link: CANDIDATE.linkedin,
        children: [r("LinkedIn", { s: 18, c: BLUE, u: {} })],
      }),
      r("  |  ", { s: 18, c: LGRAY }),
      new ExternalHyperlink({
        link: CANDIDATE.portfolio,
        children: [r("Portfolio", { s: 18, c: BLUE, u: {} })],
      }),
    ],
  }),
);
 
// SUMMARY
children.push(heading("Summary"));
children.push(
  new Paragraph({
    spacing: { before: 40, after: 0 },
    children: [r(CANDIDATE.summary, { s: 18, c: GRAY })],
  }),
);
 
// TECHNICAL SKILLS
children.push(heading("Technical Skills"));
SKILLS.forEach((sk) => children.push(skillRow(sk.label, sk.value)));
 
// EXPERIENCE
children.push(heading("Experience"));
EXPERIENCES.forEach((exp, idx) => {
  const ref = `b${idx}`;
  const companyLine = exp.location
    ? `${exp.company}  —  ${exp.location}`
    : exp.company;
  children.push(roleHeader(exp.title, companyLine, exp.dates));
  children.push(roleSubtitle(exp.title));
  exp.bullets.forEach((b) => children.push(bullet(b, ref)));
});
 
// PROJECTS
children.push(heading("Projects"));
PROJECTS.forEach((proj, idx) => {
  const ref = `b${EXPERIENCES.length + idx}`;
  children.push(projHeader(proj.title, proj.demo, proj.github));
  children.push(bullet(proj.bullet, ref));
});
 
// EDUCATION
children.push(heading("Education"));
EDUCATION.forEach((ed) => {
  children.push(
    new Paragraph({
      spacing: { before: 90, after: 28 },
      tabStops: [{ type: TabStopType.RIGHT, position: 9360 }],
      children: [
        r(ed.institution, { b: true, s: 20 }),
        r(`\t${ed.dates}`, { s: 18, c: LGRAY }),
      ],
    }),
  );
  children.push(
    new Paragraph({
      spacing: { before: 0, after: 0 },
      children: [r(ed.qualification, { i: true, s: 18, c: GRAY })],
    }),
  );
});
 
// ============================================================
// PACK AND SAVE
// ============================================================
 
const doc = new Document({
  numbering: { config: bRefs },
  sections: [
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 864, right: 1080, bottom: 864, left: 1080 },
        },
      },
      children,
    },
  ],
});
 
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(CANDIDATE.outputFile, buf);
  console.log(`CV saved: ${CANDIDATE.outputFile}`);
});