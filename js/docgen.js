/**
 * docgen.js
 * Regenerates the "Fund Approval for Procurement" form
 * (IIT Jodhpur, Office of Research and Development, Form No.: Acct/R&D-01)
 *
 * Usage:
 *   npm install docx
 *   node docgen.js
 * Produces: Fund_Approval_for_Procurement.docx
 */

const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, VerticalAlign, ShadingType,
  HeadingLevel, LevelFormat, Tab, TabStopType, TabStopPosition,
} = require("docx");

// ---------- Shared style helpers ----------
const FONT = "Book Antiqua";
const SIZE = 22; // half-points -> 11pt

const border = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const cellBorders = { top: border, bottom: border, left: border, right: border };

function run(text, opts = {}) {
  return new TextRun({ text, font: FONT, size: SIZE, bold: !!opts.bold });
}

function para(children, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.JUSTIFIED,
    children: Array.isArray(children) ? children : [children],
  });
}

function emptyPara() {
  return new Paragraph({ children: [new TextRun({ text: "", font: FONT, size: SIZE })] });
}

function cell(children, opts = {}) {
  return new TableCell({
    borders: cellBorders,
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    columnSpan: opts.colSpan,
    rowSpan: opts.rowSpan,
    verticalAlign: opts.vAlign || VerticalAlign.TOP,
    margins: { top: 40, bottom: 40, left: 100, right: 100 },
    children: Array.isArray(children) ? children : [children],
  });
}

// Column widths (DXA), total = 9360
const COL = { num: 495, label: 2655, label2: 3015, qty: 1095, cost: 2100 };
const TABLE_WIDTH = 9360;
const COLUMN_WIDTHS = [COL.num, COL.label, COL.label2, COL.qty, COL.cost];

// ---------- Header lines ----------
const headerLines = [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [run("Indian Institute of Technology Jodhpur", { bold: true })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [run("Office of Research and Development", { bold: true })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [run("Form No.: Acct/R&D-01", { bold: true })],
  }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [run("")] }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [run("Fund Approval for Procurement", { bold: true })],
  }),
];

// ---------- Rows 1-3: Name of PI / Project Title / R&D Project No. ----------
function infoRow(num, label) {
  return new TableRow({
    children: [
      cell(para(run(String(num))), { width: COL.num }),
      cell(para(run(label)), { width: COL.label }),
      cell(emptyPara(), { width: COL.label2 + COL.qty + COL.cost, colSpan: 3 }),
    ],
  });
}

// ---------- "Lists of Items" header row ----------
const listsOfItemsRow = new TableRow({
  children: [
    cell(para(run("Lists of Items", { bold: true })), { colSpan: 5, width: TABLE_WIDTH }),
  ],
});

// ---------- NR section header row ----------
function nrSectionLabelCell(rowSpan) {
  return cell(
    [
      emptyPara(), emptyPara(), emptyPara(),
      para([run("Non Recurring Items ", {}), run("(NR)", { bold: true })], { align: AlignmentType.CENTER }),
      emptyPara(),
    ],
    { width: COL.num + COL.label, colSpan: 2, rowSpan, vAlign: VerticalAlign.TOP }
  );
}

const nrHeaderRow = new TableRow({
  children: [
    nrSectionLabelCell(5), // spans this header row + 3 blank rows + total row
    cell(para(run("Item Name ")), { width: COL.label2 }),
    cell(para(run("Quantity ")), { width: COL.qty }),
    cell(para(run("Estimated Cost (Inclusive of all taxes and charges)")), { width: COL.cost }),
  ],
});

function blankItemRow() {
  return new TableRow({
    children: [
      cell(emptyPara(), { width: COL.label2 }),
      cell(emptyPara(), { width: COL.qty }),
      cell(emptyPara(), { width: COL.cost }),
    ],
  });
}

const nrTotalRow = new TableRow({
  children: [
    cell(para(run("Total (NR)", { bold: true })), { width: COL.label2 + COL.qty, colSpan: 2 }),
    cell(emptyPara(), { width: COL.cost }),
  ],
});

// ---------- Recurring section ----------
function rSectionLabelCell(rowSpan) {
  return cell(
    [
      para([run("Recurring Items", { bold: false }), run(" (R)", { bold: true })], { align: AlignmentType.CENTER }),
      para(run("{Consumable/", { bold: true }), { align: AlignmentType.CENTER }),
      para(run("Contingency/Travel/", { bold: true }), { align: AlignmentType.CENTER }),
      para(run("Miscellaneous/", { bold: true }), { align: AlignmentType.CENTER }),
      para(run("Any other please specify}", { bold: true }), { align: AlignmentType.CENTER }),
    ],
    { width: COL.num + COL.label, colSpan: 2, rowSpan, vAlign: VerticalAlign.CENTER }
  );
}

const rHeaderRow = new TableRow({
  children: [
    rSectionLabelCell(5),
    cell(para(run("Recurring items of the project")), { width: COL.label2 }),
    cell(emptyPara(), { width: COL.qty }),
    cell(emptyPara(), { width: COL.cost }),
  ],
});

const rTotalRow = new TableRow({
  children: [
    cell(para(run("Total (R)", { bold: true })), { width: COL.label2 + COL.qty, colSpan: 2 }),
    cell(emptyPara(), { width: COL.cost }),
  ],
});

const totalNRRRow = new TableRow({
  children: [
    cell(para(run("Total (NR+R)", { bold: true })), { width: COL.num + COL.label + COL.label2 + COL.qty, colSpan: 4 }),
    cell(emptyPara(), { width: COL.cost }),
  ],
});

// ---------- Committee composition / Declaration / Note block ----------
const committeeParas = [
  para(run("Committee composition:", { bold: true })),
  para(run("Procurement for items costing above Rs. 2,00,000/- (below committee)")),
  para(run("Member I:")),
  para(run("Member II:")),
  para(run("Member III:")),
  para([
    run("Procurement for items costing above Rs. 2.5 Lakhs", { bold: true }),
    run(" (the following members in addition to above three members as mentioned in A)"),
  ]),
  para(run("AR/ DR SPS")),
  para(run("AR/ DR R&D")),
  emptyPara(),
  para(run("Declaration:", { bold: true })),
  para(run("a) I have taken consent from all members of PFC. ")),
  para(run("b) Proposed items of procurement are in line with the sanctioned order of the funding agency.  ")),
  para(run("c) Procurements will be made as per the Institute procurement rules.")),
  emptyPara(),
  para([
    run("Note:", { bold: true }),
    run(" As per Circular no. IITJ/RIG/2024-25/270 dated 12 February 2025, Pre Audit is required in the following purchase cases : "),
  ]),
  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    numbering: { reference: "noteBullets", level: 0 },
    children: [run("Pre-Audit of Purchase file above Rs. 10 Lakh;")],
  }),
  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    numbering: { reference: "noteBullets", level: 0 },
    children: [run("Pre-Audit of purchase files relating to Proprietary Article Certificate (PAC) / Single Tender Enquiry, irrespective of amount.")],
  }),
  emptyPara(),
];

const committeeRow = new TableRow({
  children: [cell(committeeParas, { colSpan: 5, width: TABLE_WIDTH })],
});

const submittedRow = new TableRow({
  children: [
    cell(para(run("Submitted for Fund Approval for procurement of items as mentioned above.                        ")), {
      colSpan: 5,
      width: TABLE_WIDTH,
    }),
  ],
});

// ---------- Assemble table ----------
const mainTable = new Table({
  width: { size: TABLE_WIDTH, type: WidthType.DXA },
  columnWidths: COLUMN_WIDTHS,
  rows: [
    infoRow(1, "Name of PI"),
    infoRow(2, "Project Title"),
    infoRow(3, "R&D Project No."),
    listsOfItemsRow,
    nrHeaderRow,
    blankItemRow(),
    blankItemRow(),
    blankItemRow(),
    nrTotalRow,
    rHeaderRow,
    blankItemRow(),
    blankItemRow(),
    blankItemRow(),
    rTotalRow,
    totalNRRRow,
    committeeRow,
    submittedRow,
  ],
});

// ---------- Document ----------
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "noteBullets",
        levels: [
          {
            level: 0,
            format: LevelFormat.DECIMAL,
            text: "%1.",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          },
        ],
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720 },
        },
      },
      children: [...headerLines, mainTable, emptyPara()],
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("Fund_Approval_for_Procurement.docx", buffer);
  console.log("Done: Fund_Approval_for_Procurement.docx");
});
