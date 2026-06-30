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

  // ---------- 1. FUND APPROVAL (replica of IIT Jodhpur Acct/R&D-01) ----------
  function buildFundApproval(meta, components) {
    const nrItems = components.filter(c => c.itemType !== 'Recurring');
    const rItems = components.filter(c => c.itemType === 'Recurring');
    const totalNR = nrItems.reduce((s, c) => s + c.qty * c.unitPrice, 0);
    const totalR = rItems.reduce((s, c) => s + c.qty * c.unitPrice, 0);

    const colW = [2200, 4000, 1400, 2900]; // sums to 10500
    const tableWidth = colW.reduce((a, b) => a + b, 0);

    function infoRow(num, label, value) {
      return new D.TableRow({
        children: [
          cell(num, { width: 400, align: D.AlignmentType.CENTER }),
          cell(label, { width: 3100, bold: true }),
          new D.TableCell({
            borders: cellBorders,
            width: { size: colW[1] + colW[2] + colW[3] - 3100 + colW[0] - 400, type: D.WidthType.DXA },
            columnSpan: 3,
            margins: { top: 60, bottom: 60, left: 100, right: 100 },
            children: [new D.Paragraph({ children: [new D.TextRun({ text: value || '' })] })]
          })
        ]
      });
    }

    const itemRows = (list) => {
      const rows = [];
      const minRows = Math.max(list.length, 3);
      for (let i = 0; i < minRows; i++) {
        const it = list[i];
        rows.push(new D.TableRow({
          children: [
            cell(it ? it.name : '', { width: colW[1] }),
            cell(it ? it.qty : '', { width: colW[2], align: D.AlignmentType.CENTER }),
            cell(it ? fmtRs(it.qty * it.unitPrice) : '', { width: colW[3], align: D.AlignmentType.RIGHT })
          ]
        }));
      }
      return rows;
    };

    const headerLines = [
      new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [new D.TextRun({ text: 'Indian Institute of Technology Jodhpur', bold: true, size: 24 })] }),
      new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [new D.TextRun({ text: 'Office of Research and Development', bold: true, size: 24 })] }),
      new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [new D.TextRun({ text: 'Form No.: Acct/R&D-01', bold: true, size: 24 })] }),
      new D.Paragraph({ text: '', spacing: { after: 100 } }),
      new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [new D.TextRun({ text: 'Fund Approval for Procurement', bold: true, size: 26 })] }),
      new D.Paragraph({ alignment: D.AlignmentType.RIGHT, children: [new D.TextRun({ text: `Ref No: ${refNo('FA')}   |   Date: ${todayStr()}`, size: 18, italics: true })], spacing: { after: 150 } }),
    ];

    const infoTable = new D.Table({
      width: { size: tableWidth, type: D.WidthType.DXA },
      columnWidths: [400, 3100, 7000],
      rows: [
        infoRow(1, 'Name of PI', meta.piName),
        infoRow(2, 'Project Title', meta.projectTitle),
        infoRow(3, 'R&D Project No.', meta.projectNo)
      ]
    });

    const listHeader = new D.Paragraph({
      shading: { fill: 'E8E8E8', type: D.ShadingType.CLEAR },
      children: [new D.TextRun({ text: 'Lists of Items', bold: true })],
      spacing: { before: 100 }
    });

    const itemsHeaderRow = new D.TableRow({
      children: [
        cell('Item Name', { width: colW[1], bold: true, shading: 'F0F0F0' }),
        cell('Quantity', { width: colW[2], bold: true, shading: 'F0F0F0', align: D.AlignmentType.CENTER }),
        cell('Estimated Cost (incl. all taxes)', { width: colW[3], bold: true, shading: 'F0F0F0', align: D.AlignmentType.CENTER })
      ]
    });

    const nrTable = new D.Table({
      width: { size: tableWidth, type: D.WidthType.DXA },
      columnWidths: [colW[1], colW[2], colW[3]],
      rows: [
        new D.TableRow({ children: [cell('Non-Recurring Items (NR)', { width: colW[1], bold: true }), cell('', { width: colW[2] }), cell('', { width: colW[3] })] }),
        itemsHeaderRow,
        ...itemRows(nrItems),
        new D.TableRow({ children: [cell('Total (NR)', { width: colW[1], bold: true }), cell('', { width: colW[2] }), cell(fmtRs(totalNR), { width: colW[3], bold: true, align: D.AlignmentType.RIGHT })] })
      ]
    });

    const rTable = new D.Table({
      width: { size: tableWidth, type: D.WidthType.DXA },
      columnWidths: [colW[1], colW[2], colW[3]],
      rows: [
        new D.TableRow({ children: [cell('Recurring Items (R) — Consumable / Contingency / Travel / Misc.', { width: colW[1], bold: true }), cell('', { width: colW[2] }), cell('', { width: colW[3] })] }),
        itemsHeaderRow,
        ...itemRows(rItems),
        new D.TableRow({ children: [cell('Total (R)', { width: colW[1], bold: true }), cell('', { width: colW[2] }), cell(fmtRs(totalR), { width: colW[3], bold: true, align: D.AlignmentType.RIGHT })] })
      ]
    });

    const grandTotalTable = new D.Table({
      width: { size: tableWidth, type: D.WidthType.DXA },
      columnWidths: [colW[1] + colW[2], colW[3]],
      rows: [new D.TableRow({ children: [cell('Total (NR + R)', { width: colW[1] + colW[2], bold: true }), cell(fmtRs(totalNR + totalR), { width: colW[3], bold: true, align: D.AlignmentType.RIGHT, size: 22 })] })]
    });

    const committeeText = [
      new D.Paragraph({ text: '', spacing: { before: 150 } }),
      new D.Paragraph({ children: [new D.TextRun({ text: 'Committee composition:', bold: true })] }),
      new D.Paragraph({ text: 'A. Procurement for items costing above Rs. 2,00,000/- (below committee)', indent: { left: 300 } }),
      new D.Paragraph({ text: 'i)  Member I: _______________', indent: { left: 600 } }),
      new D.Paragraph({ text: 'ii) Member II: _______________', indent: { left: 600 } }),
      new D.Paragraph({ text: 'iii) Member III: _______________', indent: { left: 600 } }),
      new D.Paragraph({ text: '', spacing: { before: 100 } }),
      new D.Paragraph({ children: [new D.TextRun({ text: 'B. Procurement for items costing above Rs. 2.5 Lakhs', bold: true }), new D.TextRun({ text: ' (the following members in addition to the above three members in A)' })], indent: { left: 300 } }),
      new D.Paragraph({ text: 'iv) AR/ DR SPS: _______________', indent: { left: 600 } }),
      new D.Paragraph({ text: 'v)  AR/ DR R&D: _______________', indent: { left: 600 } }),
      new D.Paragraph({ text: '', spacing: { before: 150 } }),
      new D.Paragraph({ children: [new D.TextRun({ text: 'Declaration:', bold: true })] }),
      new D.Paragraph({ text: 'a) I have taken consent from all members of PFC.' }),
      new D.Paragraph({ text: 'b) Proposed items of procurement are in line with the sanctioned order of the funding agency.' }),
      new D.Paragraph({ text: 'c) Procurements will be made as per the Institute procurement rules.' }),
      new D.Paragraph({ text: '', spacing: { before: 150 } }),
      new D.Paragraph({ children: [new D.TextRun({ text: 'Note: ', bold: true }), new D.TextRun({ text: 'As per Circular no. IITJ/RIG/2024-25/270 dated 12 February 2025, Pre-Audit is required in the following purchase cases:' })] }),
      new D.Paragraph({ text: '1. Pre-Audit of Purchase file above Rs. 10 Lakh.', indent: { left: 300 } }),
      new D.Paragraph({ text: '2. Pre-Audit of purchase files relating to Proprietary Article Certificate (PAC) / Single Tender Enquiry, irrespective of amount.', indent: { left: 300 } }),
      new D.Paragraph({ text: '', spacing: { before: 200 } }),
      new D.Paragraph({ children: [new D.TextRun({ text: 'Submitted for Fund Approval for procurement of items as mentioned above.', italics: true })] }),
      new D.Paragraph({ text: '', spacing: { before: 400 } }),
    ];

    const sigTable = new D.Table({
      width: { size: tableWidth, type: D.WidthType.DXA },
      columnWidths: [3500, 3500, 3500],
      rows: [new D.TableRow({
        children: [
          new D.TableCell({ borders: { top: border, bottom: { style: D.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, left: { style: D.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, right: { style: D.BorderStyle.NONE, size: 0, color: 'FFFFFF' } }, width: { size: 3500, type: D.WidthType.DXA }, margins: { top: 100 }, children: [new D.Paragraph({ text: 'Requested by (PI)' }), new D.Paragraph({ text: '' }), new D.Paragraph({ text: 'Sign / Date: _____________' })] }),
          new D.TableCell({ borders: { top: border, bottom: { style: D.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, left: { style: D.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, right: { style: D.BorderStyle.NONE, size: 0, color: 'FFFFFF' } }, width: { size: 3500, type: D.WidthType.DXA }, margins: { top: 100 }, children: [new D.Paragraph({ text: 'HOD Approval' }), new D.Paragraph({ text: '' }), new D.Paragraph({ text: 'Sign / Date: _____________' })] }),
          new D.TableCell({ borders: { top: border, bottom: { style: D.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, left: { style: D.BorderStyle.NONE, size: 0, color: 'FFFFFF' }, right: { style: D.BorderStyle.NONE, size: 0, color: 'FFFFFF' } }, width: { size: 3500, type: D.WidthType.DXA }, margins: { top: 100 }, children: [new D.Paragraph({ text: 'Office of R&D' }), new D.Paragraph({ text: '' }), new D.Paragraph({ text: 'Sign / Date: _____________' })] }),
        ]
      })]
    });

    return new D.Document({
      styles: { default: { document: { run: { font: 'Arial', size: 20 } } } },
      sections: [{
        properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1000, bottom: 1000, left: 1000, right: 1000 } } },
        children: [...headerLines, infoTable, listHeader, nrTable, rTable, grandTotalTable, ...committeeText, sigTable]
      }]
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
