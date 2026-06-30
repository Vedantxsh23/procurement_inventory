/* WeRoCon Lab — Procurement & Inventory System
   Document generator. Builds .docx files in-browser using the vendored
   docx.js (window.docx) library — no server required.
*/

const DocGen = (() => {
  const D = window.docx;

  const border = { style: D.BorderStyle.SINGLE, size: 4, color: '000000' };
  const cellBorders = { top: border, bottom: border, left: border, right: border };

  function cell(text, opts = {}) {
    const { bold = false, width = 1000, align = D.AlignmentType.LEFT, shading = null, size = 20 } = opts;
    return new D.TableCell({
      borders: cellBorders,
      width: { size: width, type: D.WidthType.DXA },
      shading: shading ? { fill: shading, type: D.ShadingType.CLEAR } : undefined,
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [new D.Paragraph({
        alignment: align,
        children: [new D.TextRun({ text: String(text ?? ''), bold, size })]
      })]
    });
  }

  function refNo(prefix) {
    const n = Math.floor(Math.random() * 899) + 100;
    return `${prefix}-${new Date().getFullYear()}-${String(n).padStart(3, '0')}`;
  }

  function todayStr() {
    return new Date().toLocaleDateString('en-IN', { day: '2-digit', month: 'long', year: 'numeric' });
  }

  function fmtRs(n) {
    return 'Rs. ' + Number(n || 0).toLocaleString('en-IN');
  }

  // ---------- 1. FUND APPROVAL (exact replica of IIT Jodhpur Acct/R&D-01) ----------
  function buildFundApproval(meta, components) {
    const FONT = 'Book Antiqua';
    const SIZE = 22; // 11pt

    const nrItems = components.filter(c => c.itemType !== 'Recurring');
    const rItems = components.filter(c => c.itemType === 'Recurring');
    const totalNR = nrItems.reduce((s, c) => s + c.qty * c.unitPrice, 0);
    const totalR = rItems.reduce((s, c) => s + c.qty * c.unitPrice, 0);

    // Column widths (DXA), total = 9360 — matches the original form's grid
    const COL = { num: 495, label: 2655, label2: 3015, qty: 1095, cost: 2100 };
    const TABLE_WIDTH = 9360;
    const COLUMN_WIDTHS = [COL.num, COL.label, COL.label2, COL.qty, COL.cost];

    function run(text, opts = {}) {
      return new D.TextRun({ text: String(text ?? ''), font: FONT, size: SIZE, bold: !!opts.bold });
    }
    function para(children, opts = {}) {
      return new D.Paragraph({
        alignment: opts.align || D.AlignmentType.JUSTIFIED,
        children: Array.isArray(children) ? children : [children],
      });
    }
    function emptyPara() {
      return new D.Paragraph({ children: [new D.TextRun({ text: '', font: FONT, size: SIZE })] });
    }
    function fcell(children, opts = {}) {
      return new D.TableCell({
        borders: cellBorders,
        width: opts.width ? { size: opts.width, type: D.WidthType.DXA } : undefined,
        columnSpan: opts.colSpan,
        rowSpan: opts.rowSpan,
        verticalAlign: opts.vAlign || D.VerticalAlign.TOP,
        margins: { top: 40, bottom: 40, left: 100, right: 100 },
        children: Array.isArray(children) ? children : [children],
      });
    }

    // ---- Header lines ----
    const headerLines = [
      new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [run('Indian Institute of Technology Jodhpur', { bold: true })] }),
      new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [run('Office of Research and Development', { bold: true })] }),
      new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [run('Form No.: Acct/R&D-01', { bold: true })] }),
      new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [run('')] }),
      new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [run('Fund Approval for Procurement', { bold: true })] }),
    ];

    // ---- Rows 1-3: Name of PI / Project Title / R&D Project No. ----
    function infoRow(num, label, value) {
      return new D.TableRow({
        children: [
          fcell(para(run(String(num))), { width: COL.num }),
          fcell(para(run(label)), { width: COL.label }),
          fcell(para(run(value || '')), { width: COL.label2 + COL.qty + COL.cost, colSpan: 3 }),
        ],
      });
    }

    // ---- "Lists of Items" header row ----
    const listsOfItemsRow = new D.TableRow({
      children: [fcell(para(run('Lists of Items', { bold: true })), { colSpan: 5, width: TABLE_WIDTH })],
    });

    // ---- NR section ----
    const nrRowCount = Math.max(nrItems.length, 3);
    function nrSectionLabelCell() {
      return fcell(
        [
          emptyPara(), emptyPara(), emptyPara(),
          para([run('Non Recurring Items '), run('(NR)', { bold: true })], { align: D.AlignmentType.CENTER }),
          emptyPara(),
        ],
        { width: COL.num + COL.label, colSpan: 2, rowSpan: nrRowCount + 2, vAlign: D.VerticalAlign.TOP }
      );
    }
    const nrHeaderRow = new D.TableRow({
      children: [
        nrSectionLabelCell(),
        fcell(para(run('Item Name ')), { width: COL.label2 }),
        fcell(para(run('Quantity ')), { width: COL.qty }),
        fcell(para(run('Estimated Cost (Inclusive of all taxes and charges)')), { width: COL.cost }),
      ],
    });
    function itemRow(it) {
      return new D.TableRow({
        children: [
          fcell(para(run(it ? it.name : '')), { width: COL.label2 }),
          fcell(para(run(it ? String(it.qty) : ''), { align: D.AlignmentType.CENTER }), { width: COL.qty }),
          fcell(para(run(it ? fmtRs(it.qty * it.unitPrice) : ''), { align: D.AlignmentType.RIGHT }), { width: COL.cost }),
        ],
      });
    }
    const nrItemRows = [];
    for (let i = 0; i < nrRowCount; i++) nrItemRows.push(itemRow(nrItems[i]));
    const nrTotalRow = new D.TableRow({
      children: [
        fcell(para(run('Total (NR)', { bold: true })), { width: COL.label2 + COL.qty, colSpan: 2 }),
        fcell(para(run(fmtRs(totalNR)), { align: D.AlignmentType.RIGHT }), { width: COL.cost }),
      ],
    });

    // ---- Recurring section ----
    const rRowCount = Math.max(rItems.length, 3);
    function rSectionLabelCell() {
      return fcell(
        [
          para([run('Recurring Items'), run(' (R)', { bold: true })], { align: D.AlignmentType.CENTER }),
          para(run('{Consumable/', { bold: true }), { align: D.AlignmentType.CENTER }),
          para(run('Contingency/Travel/', { bold: true }), { align: D.AlignmentType.CENTER }),
          para(run('Miscellaneous/', { bold: true }), { align: D.AlignmentType.CENTER }),
          para(run('Any other please specify}', { bold: true }), { align: D.AlignmentType.CENTER }),
        ],
        { width: COL.num + COL.label, colSpan: 2, rowSpan: rRowCount + 2, vAlign: D.VerticalAlign.CENTER }
      );
    }
    const rHeaderRow = new D.TableRow({
      children: [
        rSectionLabelCell(),
        fcell(para(run('Recurring items of the project')), { width: COL.label2 }),
        fcell(emptyPara(), { width: COL.qty }),
        fcell(emptyPara(), { width: COL.cost }),
      ],
    });
    const rItemRows = [];
    for (let i = 0; i < rRowCount; i++) rItemRows.push(itemRow(rItems[i]));
    const rTotalRow = new D.TableRow({
      children: [
        fcell(para(run('Total (R)', { bold: true })), { width: COL.label2 + COL.qty, colSpan: 2 }),
        fcell(para(run(fmtRs(totalR)), { align: D.AlignmentType.RIGHT }), { width: COL.cost }),
      ],
    });

    const totalNRRRow = new D.TableRow({
      children: [
        fcell(para(run('Total (NR+R)', { bold: true })), { width: COL.num + COL.label + COL.label2 + COL.qty, colSpan: 4 }),
        fcell(para(run(fmtRs(totalNR + totalR)), { align: D.AlignmentType.RIGHT }), { width: COL.cost }),
      ],
    });

    // ---- Committee composition / Declaration / Note block ----
    const committeeParas = [
      para(run('Committee composition:', { bold: true })),
      para(run('Procurement for items costing above Rs. 2,00,000/- (below committee)')),
      para(run('Member I:')),
      para(run('Member II:')),
      para(run('Member III:')),
      para([
        run('Procurement for items costing above Rs. 2.5 Lakhs', { bold: true }),
        run(' (the following members in addition to above three members as mentioned in A)'),
      ]),
      para(run('AR/ DR SPS')),
      para(run('AR/ DR R&D')),
      emptyPara(),
      para(run('Declaration:', { bold: true })),
      para(run('a) I have taken consent from all members of PFC. ')),
      para(run('b) Proposed items of procurement are in line with the sanctioned order of the funding agency.  ')),
      para(run('c) Procurements will be made as per the Institute procurement rules.')),
      emptyPara(),
      para([
        run('Note:', { bold: true }),
        run(' As per Circular no. IITJ/RIG/2024-25/270 dated 12 February 2025, Pre Audit is required in the following purchase cases : '),
      ]),
      new D.Paragraph({
        alignment: D.AlignmentType.JUSTIFIED,
        numbering: { reference: 'noteBullets', level: 0 },
        children: [run('Pre-Audit of Purchase file above Rs. 10 Lakh;')],
      }),
      new D.Paragraph({
        alignment: D.AlignmentType.JUSTIFIED,
        numbering: { reference: 'noteBullets', level: 0 },
        children: [run('Pre-Audit of purchase files relating to Proprietary Article Certificate (PAC) / Single Tender Enquiry, irrespective of amount.')],
      }),
      emptyPara(),
    ];
    const committeeRow = new D.TableRow({
      children: [fcell(committeeParas, { colSpan: 5, width: TABLE_WIDTH })],
    });

    const submittedRow = new D.TableRow({
      children: [
        fcell(para(run('Submitted for Fund Approval for procurement of items as mentioned above.                        ')), {
          colSpan: 5,
          width: TABLE_WIDTH,
        }),
      ],
    });

    const mainTable = new D.Table({
      width: { size: TABLE_WIDTH, type: D.WidthType.DXA },
      columnWidths: COLUMN_WIDTHS,
      rows: [
        infoRow(1, 'Name of PI', meta.piName),
        infoRow(2, 'Project Title', meta.projectTitle),
        infoRow(3, 'R&D Project No.', meta.projectNo),
        listsOfItemsRow,
        nrHeaderRow,
        ...nrItemRows,
        nrTotalRow,
        rHeaderRow,
        ...rItemRows,
        rTotalRow,
        totalNRRRow,
        committeeRow,
        submittedRow,
      ],
    });

    return new D.Document({
      numbering: {
        config: [
          {
            reference: 'noteBullets',
            levels: [
              {
                level: 0,
                format: D.LevelFormat.DECIMAL,
                text: '%1.',
                alignment: D.AlignmentType.LEFT,
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
  }

  // ---------- 2. QUOTATION COMPARISON ----------
  function buildQuotationDoc(meta, components) {
    const colW = [600, 3200, 1000, 2400, 2400];
    const tableWidth = colW.reduce((a, b) => a + b, 0);
    const headerRow = new D.TableRow({
      children: [
        cell('S.No', { width: colW[0], bold: true, shading: 'F0F0F0', align: D.AlignmentType.CENTER }),
        cell('Component', { width: colW[1], bold: true, shading: 'F0F0F0' }),
        cell('Qty', { width: colW[2], bold: true, shading: 'F0F0F0', align: D.AlignmentType.CENTER }),
        cell('Selected vendor', { width: colW[3], bold: true, shading: 'F0F0F0' }),
        cell('Quoted amount', { width: colW[4], bold: true, shading: 'F0F0F0', align: D.AlignmentType.RIGHT })
      ]
    });
    const rows = components.map((c, i) => new D.TableRow({
      children: [
        cell(i + 1, { width: colW[0], align: D.AlignmentType.CENTER }),
        cell(c.name, { width: colW[1] }),
        cell(c.qty, { width: colW[2], align: D.AlignmentType.CENTER }),
        cell(c.vendor || '-', { width: colW[3] }),
        cell(fmtRs(c.qty * c.unitPrice), { width: colW[4], align: D.AlignmentType.RIGHT })
      ]
    }));
    const total = components.reduce((s, c) => s + c.qty * c.unitPrice, 0);
    const totalRow = new D.TableRow({
      children: [
        cell('', { width: colW[0] }),
        cell('', { width: colW[1] }),
        cell('', { width: colW[2] }),
        cell('Total', { width: colW[3], bold: true }),
        cell(fmtRs(total), { width: colW[4], bold: true, align: D.AlignmentType.RIGHT })
      ]
    });

    return new D.Document({
      styles: { default: { document: { run: { font: 'Arial', size: 20 } } } },
      sections: [{
        properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1000, bottom: 1000, left: 1000, right: 1000 } } },
        children: [
          new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [new D.TextRun({ text: meta.labName || 'WeRoCon Lab', bold: true, size: 24 })] }),
          new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [new D.TextRun({ text: 'Vendor Quotation Comparison Statement', bold: true, size: 26 })], spacing: { after: 100 } }),
          new D.Paragraph({ alignment: D.AlignmentType.RIGHT, children: [new D.TextRun({ text: `Ref No: ${refNo('QT')}   |   Date: ${todayStr()}`, italics: true, size: 18 })], spacing: { after: 200 } }),
          new D.Paragraph({ children: [new D.TextRun({ text: `Project: ${meta.projectTitle || '-'}   |   PI: ${meta.piName || '-'}`, size: 18 })], spacing: { after: 200 } }),
          new D.Table({ width: { size: tableWidth, type: D.WidthType.DXA }, columnWidths: colW, rows: [headerRow, ...rows, totalRow] }),
          new D.Paragraph({ text: '', spacing: { before: 200 } }),
          new D.Paragraph({ text: 'Certified that minimum 3 quotations were obtained per item as per procurement policy. Quotation copies attached separately.', italics: true, size: 18 }),
          new D.Paragraph({ text: '', spacing: { before: 400 } }),
          new D.Paragraph({ children: [new D.TextRun({ text: 'Prepared by: _____________      Approved by (HOD): _____________' })] })
        ]
      }]
    });
  }

  // ---------- 3. NON-GeM CERTIFICATE ----------
  function buildNonGemDoc(meta, components) {
    const nonGemItems = components.filter(c => c.gemStatus === 'Non-GeM certified' || c.gemStatus === 'Not checked');
    const colW = [600, 3400, 3000, 3000];
    const tableWidth = colW.reduce((a, b) => a + b, 0);
    const headerRow = new D.TableRow({
      children: [
        cell('S.No', { width: colW[0], bold: true, shading: 'F0F0F0', align: D.AlignmentType.CENTER }),
        cell('Item description', { width: colW[1], bold: true, shading: 'F0F0F0' }),
        cell('GeM search reference', { width: colW[2], bold: true, shading: 'F0F0F0' }),
        cell('Reason for non-GeM', { width: colW[3], bold: true, shading: 'F0F0F0' })
      ]
    });
    const rows = nonGemItems.map((c, i) => new D.TableRow({
      children: [
        cell(i + 1, { width: colW[0], align: D.AlignmentType.CENTER }),
        cell(c.name, { width: colW[1] }),
        cell(c.gemSearchRef || `GEM/SRCH/${new Date().getFullYear()}/${String(i + 1).padStart(4, '0')}`, { width: colW[2] }),
        cell('Not available meeting required technical specifications on GeM portal', { width: colW[3], size: 18 })
      ]
    }));

    return new D.Document({
      styles: { default: { document: { run: { font: 'Arial', size: 20 } } } },
      sections: [{
        properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1000, bottom: 1000, left: 1000, right: 1000 } } },
        children: [
          new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [new D.TextRun({ text: meta.labName || 'WeRoCon Lab', bold: true, size: 24 })] }),
          new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [new D.TextRun({ text: 'Certificate for Non-GeM Purchase', bold: true, size: 26 })] }),
          new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [new D.TextRun({ text: '(As per GFR 2017 Rule 149 / GeM procurement guidelines)', italics: true, size: 18 })], spacing: { after: 100 } }),
          new D.Paragraph({ alignment: D.AlignmentType.RIGHT, children: [new D.TextRun({ text: `Cert No: ${refNo('NG')}   |   Date: ${todayStr()}`, italics: true, size: 18 })], spacing: { after: 200 } }),
          new D.Paragraph({ text: 'This is to certify that the following items were searched on the Government e-Marketplace (GeM) portal (gem.gov.in) and are either not available or do not meet the required technical specifications. Purchase is therefore proposed through the open market / direct procurement route:', spacing: { after: 150 } }),
          new D.Table({ width: { size: tableWidth, type: D.WidthType.DXA }, columnWidths: colW, rows: [headerRow, ...rows] }),
          new D.Paragraph({ text: '', spacing: { before: 300 } }),
          new D.Paragraph({ children: [new D.TextRun({ text: 'Certified by: _____________      Designation: _____________' })] }),
          new D.Paragraph({ text: '', spacing: { before: 150 } }),
          new D.Paragraph({ children: [new D.TextRun({ text: 'Countersigned (HOD): _____________      Date: _____________' })] })
        ]
      }]
    });
  }

  // ---------- 4. PAYMENT RECEIPT ----------
  function buildPaymentReceiptDoc(meta, components) {
    const invoiced = components.filter(c => c.invoiceNo);
    const colW = [600, 2800, 1800, 2200, 1900, 1700];
    const tableWidth = colW.reduce((a, b) => a + b, 0);
    const headerRow = new D.TableRow({
      children: [
        cell('S.No', { width: colW[0], bold: true, shading: 'F0F0F0', align: D.AlignmentType.CENTER }),
        cell('Component', { width: colW[1], bold: true, shading: 'F0F0F0' }),
        cell('Invoice no.', { width: colW[2], bold: true, shading: 'F0F0F0' }),
        cell('Vendor', { width: colW[3], bold: true, shading: 'F0F0F0' }),
        cell('Amount', { width: colW[4], bold: true, shading: 'F0F0F0', align: D.AlignmentType.RIGHT }),
        cell('Status', { width: colW[5], bold: true, shading: 'F0F0F0' })
      ]
    });
    const rows = invoiced.map((c, i) => new D.TableRow({
      children: [
        cell(i + 1, { width: colW[0], align: D.AlignmentType.CENTER }),
        cell(c.name, { width: colW[1] }),
        cell(c.invoiceNo, { width: colW[2] }),
        cell(c.vendor || '-', { width: colW[3] }),
        cell(fmtRs(c.qty * c.unitPrice), { width: colW[4], align: D.AlignmentType.RIGHT }),
        cell(c.paymentStatus, { width: colW[5], size: 16 })
      ]
    }));
    const total = invoiced.reduce((s, c) => s + c.qty * c.unitPrice, 0);

    return new D.Document({
      styles: { default: { document: { run: { font: 'Arial', size: 20 } } } },
      sections: [{
        properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1000, bottom: 1000, left: 1000, right: 1000 } } },
        children: [
          new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [new D.TextRun({ text: meta.labName || 'WeRoCon Lab', bold: true, size: 24 })] }),
          new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [new D.TextRun({ text: 'Payment Receipt / Reimbursement Claim', bold: true, size: 26 })] }),
          new D.Paragraph({ alignment: D.AlignmentType.RIGHT, children: [new D.TextRun({ text: `Receipt No: ${refNo('PR')}   |   Date: ${todayStr()}`, italics: true, size: 18 })], spacing: { after: 200 } }),
          new D.Table({ width: { size: tableWidth, type: D.WidthType.DXA }, columnWidths: colW, rows: rows.length ? [headerRow, ...rows] : [headerRow] }),
          new D.Paragraph({ text: '', spacing: { before: 100 } }),
          new D.Paragraph({ children: [new D.TextRun({ text: `Total: ${fmtRs(total)}`, bold: true })] }),
          new D.Paragraph({ text: '', spacing: { before: 400 } }),
          new D.Paragraph({ children: [new D.TextRun({ text: 'Claimed by: _____________      Verified by: _____________      Accounts: _____________' })] })
        ]
      }]
    });
  }

  // ---------- triggers ----------
  function triggerDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  async function generateBlob(type, meta, components) {
    let doc;
    if (type === 'fund-approval') doc = buildFundApproval(meta, components);
    else if (type === 'quotation') doc = buildQuotationDoc(meta, components);
    else if (type === 'non-gem') doc = buildNonGemDoc(meta, components);
    else if (type === 'payment-receipt') doc = buildPaymentReceiptDoc(meta, components);
    else throw new Error('Unknown doc type: ' + type);
    return await D.Packer.toBlob(doc);
  }

  async function downloadOne(type, meta, components, filename) {
    const blob = await generateBlob(type, meta, components);
    triggerDownload(blob, filename);
  }

  // Bundle all docs + any uploaded quotation/invoice files into one ZIP
  async function downloadBundle(meta, components) {
    const zip = new JSZip();
    const stamp = new Date().toISOString().split('T')[0];

    const fund = await generateBlob('fund-approval', meta, components);
    zip.file(`Fund_Approval_${stamp}.docx`, fund);

    const quote = await generateBlob('quotation', meta, components);
    zip.file(`Quotation_Comparison_${stamp}.docx`, quote);

    const nonGemItems = components.filter(c => c.gemStatus === 'Non-GeM certified' || c.gemStatus === 'Not checked');
    if (nonGemItems.length) {
      const nonGem = await generateBlob('non-gem', meta, components);
      zip.file(`Non_GeM_Certificate_${stamp}.docx`, nonGem);
    }

    const invoicedItems = components.filter(c => c.invoiceNo);
    if (invoicedItems.length) {
      const receipt = await generateBlob('payment-receipt', meta, components);
      zip.file(`Payment_Receipt_${stamp}.docx`, receipt);
    }

    // include any uploaded quotation / invoice files
    const uploadsFolder = zip.folder('Uploaded_Files');
    components.forEach(c => {
      (c.quotationFiles || []).forEach((f, idx) => {
        if (f.dataUrl) {
          const base64 = f.dataUrl.split(',')[1];
          uploadsFolder.file(`${c.name.replace(/[^a-z0-9]/gi, '_')}_quote${idx + 1}_${f.name}`, base64, { base64: true });
        }
      });
      if (c.invoiceFile && c.invoiceFile.dataUrl) {
        const base64 = c.invoiceFile.dataUrl.split(',')[1];
        uploadsFolder.file(`${c.name.replace(/[^a-z0-9]/gi, '_')}_invoice_${c.invoiceFile.name}`, base64, { base64: true });
      }
    });

    const blob = await zip.generateAsync({ type: 'blob' });
    triggerDownload(blob, `WeRoCon_Procurement_Bundle_${stamp}.zip`);
  }

  return { downloadOne, downloadBundle, generateBlob, refNo, todayStr, fmtRs };
})();
