/** @file Generate the synthetic demo DOCX with styled text, tables, lists, and preserved parts. */
const fs = require("node:fs");
const path = require("node:path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, PageNumber, HeadingLevel, WidthType, ShadingType,
  BorderStyle, LevelFormat, AlignmentType, ImageRun,
} = require("docx");

const border = { style: BorderStyle.SINGLE, size: 4, color: "D9D9D9" };
/**
 * Create a bordered fixture cell with explicit DXA width and optional header styling.
 * @param {string} text Synthetic cell content.
 * @param {number} width Cell width in twips (one twentieth of a point).
 * @param {boolean} [header=false] Use bold text and shaded background.
 * @returns {TableCell} A cell containing one plain-text paragraph.
 */
const cell = (text, width, header = false) => new TableCell({
  width: { size: width, type: WidthType.DXA },
  margins: { top: 100, bottom: 100, left: 140, right: 140 },
  borders: { top: border, bottom: border, left: border, right: border },
  shading: { type: ShadingType.CLEAR, fill: header ? "F2F2F2" : "FFFFFF" },
  children: [new Paragraph({ children: [new TextRun({ text, bold: header })] })],
});
const doc = new Document({
  creator: "Word Editing Showcase",
  title: "Northstar project proposal - synthetic demo",
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: "Arial", size: 32, bold: true, color: "000000" },
        paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 0 } },
    ],
  },
  numbering: { config: [{ reference: "scope", levels: [{
    level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 720, hanging: 360 } } },
  }] }] },
  sections: [{
    properties: {
      page: { size: { width: 12240, height: 15840 },
        margin: { top: 1200, right: 1440, bottom: 1200, left: 1440 } },
    },
    headers: { default: new Header({ children: [new Paragraph("NORTHSTAR | SYNTHETIC DEMO DOCUMENT")] }) },
    footers: { default: new Footer({ children: [new Paragraph({
      children: [new TextRun("Demo only - no customer data | Page "), new TextRun({ children: [PageNumber.CURRENT] })],
    })] }) },
    children: [
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Northstar project proposal")] }),
      new Paragraph("A fictional proposal for demonstrating safe, review-first Word document editing."),
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Executive summary")] }),
      new Paragraph({ children: [
        new TextRun("Northstar will help the team "),
        new TextRun({ text: "reduce manual reporting", bold: true }),
        new TextRun(" by providing a shared weekly delivery summary. The goal of this initiative is to make it easier for stakeholders to understand progress, identify risks earlier, and spend less time collecting status updates from different places."),
      ] }),
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Commercial terms")] }),
      new Paragraph({ children: [
        new TextRun("Payment is due within "),
        new TextRun({ text: "30", bold: true }),
        new TextRun(" days of invoice receipt. The project budget is "),
        new TextRun({ text: "$25,000", italics: true }),
        new TextRun(". All amounts are illustrative."),
      ] }),
      new Table({
        width: { size: 9360, type: WidthType.DXA }, columnWidths: [4680, 4680],
        rows: [
          new TableRow({ children: [cell("Milestone", 4680, true), cell("Target", 4680, true)] }),
          new TableRow({ children: [cell("Discovery", 4680), cell("Week 1", 4680)] }),
          new TableRow({ children: [cell("Pilot launch", 4680), cell("Week 4", 4680)] }),
        ],
      }),
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Scope")] }),
      new Paragraph({ numbering: { reference: "scope", level: 0 }, children: [new TextRun("Create a reusable weekly status template.")] }),
      new Paragraph({ numbering: { reference: "scope", level: 0 }, children: [new TextRun("Run a pilot with one fictional delivery team.")] }),
      new Paragraph({ children: [new ImageRun({
        type: "png",
        data: Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGPYKO/9HwAEuwIbJO/A4wAAAABJRU5ErkJggg==", "base64"),
        transformation: { width: 8, height: 8 },
        altText: { title: "Fixture marker", description: "Synthetic image preserved during edits", name: "marker" },
      })] }),
    ],
  }],
});

const output = path.join(__dirname, "..", "demo", "sample.docx");
fs.mkdirSync(path.dirname(output), { recursive: true });
/** Serialize the fixture and write it to the demo folder, surfacing generation failures. */
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(output, buffer);
  console.log("Created demo\\sample.docx");
}).catch(error => {
  console.error(error);
  process.exitCode = 1;
});
