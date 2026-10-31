import ExcelJS from 'exceljs';
import html2pdf from 'html2pdf.js';
import { BusinessDocument, BusinessSettings, documentFolio, documentTitle, subtotalOf } from '@/types';

// ---------- shared helpers ----------
const BLUE = 'FF00B0F0';
const BLACK = 'FF000000';
const WHITE = 'FFFFFFFF';
const thinBorder: Partial<ExcelJS.Border> = { style: 'thin', color: { argb: BLACK } };
const thickBorder: Partial<ExcelJS.Border> = { style: 'thick', color: { argb: BLACK } };
const allThin = { top: thinBorder, left: thinBorder, bottom: thinBorder, right: thinBorder };
const moneyFmt = '"$" #,##0.00';

function arrayBufferToBase64(buffer: ArrayBuffer): string {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

// ---------- Excel ----------
export async function downloadExcel(document: BusinessDocument, settings: BusinessSettings) {
  const folio = documentFolio(document);
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Hoja1', {
    properties: { defaultRowHeight: 12.75 },
    pageSetup: { paperSize: 1 as unknown as ExcelJS.PaperSize, orientation: 'portrait' },
  });

  // column widths
  sheet.columns = [
    { width: 6 },    // A
    { width: 10.4 }, // B
    { width: 9.6 },  // C
    { width: 34 },   // D
    { width: 6.6 },   // E
    { width: 13 },   // F
    { width: 17 },   // G
  ];

  // Row 8 is thinner
  sheet.getRow(8).height = 9;

  // Merge cells
  sheet.mergeCells('A1:C7');      // logo
  sheet.mergeCells('E1:G2');      // title
  sheet.mergeCells('F4:G4');      // date
  sheet.mergeCells('F5:G5');      // folio
  ['A9:B9', 'A10:B10', 'A11:B11', 'A12:B12', 'C9:D9', 'C10:D10', 'C11:D11', 'C12:D12'].forEach((r) => sheet.mergeCells(r));
  sheet.mergeCells('D14:E14');    // Descripción header
  for (let i = 15; i <= 27; i++) sheet.mergeCells(`D${i}:E${i}`);
  sheet.mergeCells('A29:D29');
  sheet.mergeCells('A30:C30');
  sheet.mergeCells('A31:C31');
  sheet.mergeCells('A32:C32');
  sheet.mergeCells('A37:G37');

  // Logo image in A1:C7
  const logoResponse = await fetch('/logo-ferreteria-garcia.jpg').catch(() => null);
  if (logoResponse && logoResponse.ok) {
    const logoBuf = await logoResponse.arrayBuffer();
    const base64 = arrayBufferToBase64(logoBuf);
    const imageId = workbook.addImage({ base64: `data:image/jpeg;base64,${base64}`, extension: 'jpeg' });
    sheet.addImage(imageId, { tl: { col: 0, row: 0 } as any, br: { col: 3, row: 7 } as any, editAs: 'oneCell' });
  }

  // E1 title
  const titleCell = sheet.getCell('E1');
  titleCell.value = documentTitle(document.type);
  titleCell.font = { name: 'Century Gothic', size: 22, bold: true, color: { argb: BLUE } };
  titleCell.alignment = { horizontal: 'right', vertical: 'middle' };

  // E4 FECHA, E5 FOLIO
  sheet.getCell('E4').value = 'FECHA:';
  sheet.getCell('E4').font = { name: 'Calibri', size: 10, bold: true };
  sheet.getCell('E4').alignment = { horizontal: 'right' };

  const dateCell = sheet.getCell('F4');
  const dt = new Date(`${document.date}T12:00:00`);
  dateCell.value = dt;
  dateCell.numFmt = 'd-mmm-yy';
  dateCell.font = { name: 'Calibri', size: 10 };
  dateCell.alignment = { horizontal: 'center' };
  dateCell.border = { top: thickBorder, bottom: thickBorder, left: thickBorder, right: thickBorder };

  sheet.getCell('E5').value = 'FOLIO:';
  sheet.getCell('E5').font = { name: 'Calibri', size: 10, bold: true };
  sheet.getCell('E5').alignment = { horizontal: 'right' };

  const folioCell = sheet.getCell('F5');
  folioCell.value = `#${folio}`;
  folioCell.font = { name: 'Calibri', size: 10 };
  folioCell.alignment = { horizontal: 'center' };
  folioCell.border = { top: thickBorder, bottom: thickBorder, left: thickBorder, right: thickBorder };

  // Row 8 thin red separator line
  for (let col = 1; col <= 7; col++) {
    const c = sheet.getRow(8).getCell(col);
    c.border = { bottom: { style: 'thin', color: { argb: 'FFFF0000' } } };
  }

  // Client info rows 9-12
  const clientLabels: [string, string][] = [
    ['A9', 'Cliente:'],
    ['A10', 'Atención:'],
    ['A11', 'Correo Electrónico'],
    ['A12', 'Teléfono:'],
  ];
  const clientValues = [document.clientName, document.attention || '', document.email || '', document.phone || ''];
  clientLabels.forEach(([ref, text], i) => {
    const cell = sheet.getCell(ref);
    cell.value = text;
    cell.font = { name: 'Calibri', size: 9, bold: true };
    cell.alignment = { horizontal: 'right' };
  });
  clientValues.forEach((val, i) => {
    const cell = sheet.getCell(`C${9 + i}`);
    cell.value = val;
    cell.font = { name: 'Calibri', size: 10 };
  });

  // Row 14 headers
  const headers: [string, string, Partial<ExcelJS.Alignment>][] = [
    ['A14', 'No.', { horizontal: 'center' }],
    ['B14', 'Cantidad', { horizontal: 'center' }],
    ['C14', 'Unidad', { horizontal: 'center' }],
    ['D14', 'Descripción', { horizontal: 'center' }],
    ['F14', 'Precio Unit.', { horizontal: 'center' }],
    ['G14', 'Total', { horizontal: 'center' }],
  ];
  headers.forEach(([ref, text, align]) => {
    const cell = sheet.getCell(ref);
    cell.value = text;
    cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: WHITE } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: BLUE } };
    cell.alignment = { ...align, vertical: 'middle' } as ExcelJS.Alignment;
    cell.border = allThin;
  });
  // E14 also gets blue fill (part of merged Descripción header area D14:E14 but E is covered)
  sheet.getCell('E14').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: BLUE } };
  sheet.getCell('E14').border = allThin;

  // Rows 15-27 item data
  const maxItems = 13; // rows 15..27
  for (let i = 0; i < maxItems; i++) {
    const row = 15 + i;
    const item = document.items[i];
    sheet.getCell(`A${row}`).value = item ? i + 1 : '';
    sheet.getCell(`A${row}`).alignment = { horizontal: 'center' };
    sheet.getCell(`B${row}`).value = item ? item.quantity : '';
    sheet.getCell(`B${row}`).alignment = { horizontal: 'center' };
    sheet.getCell(`C${row}`).value = item ? item.unit : '';
    sheet.getCell(`C${row}`).alignment = { horizontal: 'center' };
    sheet.getCell(`D${row}`).value = item ? item.description : '';

    if (item) {
      const priceCell = sheet.getCell(`F${row}`);
      priceCell.value = item.unitPrice;
      priceCell.numFmt = moneyFmt;
      const totalCell = sheet.getCell(`G${row}`);
      totalCell.value = { formula: `B${row}*F${row}` };
      totalCell.numFmt = moneyFmt;
    }

    for (const col of ['A', 'B', 'C', 'D', 'E', 'F', 'G'] as const) {
      sheet.getCell(`${col}${row}`).border = allThin;
    }
  }

  // Row 28 SUBTOTAL
  sheet.getCell('F28').value = 'SUBTOTAL';
  sheet.getCell('F28').font = { name: 'Calibri', size: 10, bold: true };
  sheet.getCell('F28').alignment = { horizontal: 'right' };
  const subCell = sheet.getCell('G28');
  subCell.value = { formula: 'SUM(G15:G27)' };
  subCell.numFmt = moneyFmt;
  subCell.font = { bold: true };
  subCell.border = allThin;
  sheet.getCell('F28').border = allThin;

  // Row 29 IVA (G29)
  sheet.getCell('F29').value = 'IVA';
  sheet.getCell('F29').font = { name: 'Calibri', size: 10, bold: true };
  sheet.getCell('F29').alignment = { horizontal: 'right' };
  const ivaCell = sheet.getCell('G29');
  ivaCell.value = { formula: 'G28*0.16' };
  ivaCell.numFmt = moneyFmt;
  ivaCell.font = { bold: true };
  ivaCell.border = allThin;
  sheet.getCell('F29').border = allThin;

  // Row 30 TOTAL (G30)
  sheet.getCell('F30').value = 'TOTAL';
  sheet.getCell('F30').font = { name: 'Calibri', size: 10, bold: true };
  sheet.getCell('F30').alignment = { horizontal: 'right' };
  const totalCell = sheet.getCell('G30');
  totalCell.value = { formula: 'G28+G29' };
  totalCell.numFmt = moneyFmt;
  totalCell.font = { bold: true };
  totalCell.border = allThin;
  sheet.getCell('F30').border = allThin;

  // A29:D29 Observaciones header
  const obsHeader = sheet.getCell('A29');
  obsHeader.value = 'Observaciones';
  obsHeader.font = { name: 'Calibri', size: 10, bold: true, color: { argb: WHITE } };
  obsHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: BLUE } };
  obsHeader.alignment = { horizontal: 'center', vertical: 'middle' };

  // A30 Tiempo de Entrega, A31 Forma de Pago, A32 Nota
  const obsRows: [string, string, string][] = [
    ['A30', 'Tiempo de Entrega:', document.delivery],
    ['A31', 'Forma de Pago:', document.payment],
    ['A32', 'Nota:', document.note],
  ];
  obsRows.forEach(([labelRef, labelText, val], idx) => {
    const labelCell = sheet.getCell(labelRef);
    labelCell.value = labelText;
    labelCell.font = { name: 'Calibri', size: 10, bold: true };
    labelCell.alignment = { horizontal: 'right' };
    const valCell = sheet.getCell(`D${30 + idx}`);
    valCell.value = val;
    valCell.font = { name: 'Calibri', size: 10 };
    if (idx === 2) {
      valCell.border = { bottom: thickBorder };
    }
  });

  // D36 Sello o firma de recibido
  const signCell = sheet.getCell('D36');
  signCell.value = 'Sello o firma de recibido';
  signCell.font = { name: 'Calibri', size: 10, bold: true };
  signCell.alignment = { horizontal: 'center' };
  signCell.border = { top: { style: 'thin', color: { argb: BLACK } } };

  // A37 footer
  const footerCell = sheet.getCell('A37');
  footerCell.value = `Ventas: ${settings.salesName}, ${settings.salesPhone}. Correo: ${settings.salesEmail}`;
  footerCell.font = { name: 'Calibri', size: 12, bold: true };
  footerCell.alignment = { horizontal: 'center', vertical: 'middle' };

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const anchor = globalThis.document.createElement('a');
  anchor.href = url;
  anchor.download = `${folio}.xlsx`;
  anchor.click();
  URL.revokeObjectURL(url);
}

// ---------- PDF (via html2pdf.js from preview HTML) ----------
export function downloadPdf(document: BusinessDocument, settings: BusinessSettings) {
  const folio = documentFolio(document);
  const html = buildPreviewHtml(document, settings);
  const container = globalThis.document.createElement('div');
  container.innerHTML = html;
  globalThis.document.body.appendChild(container);

  html2pdf()
    .set({
      margin: 0,
      filename: `${folio}.pdf`,
      image: { type: 'jpeg', quality: 0.98 },
      html2canvas: { scale: 2, useCORS: true, letterRendering: true },
      jsPDF: { unit: 'mm', format: 'letter', orientation: 'portrait' },
    })
    .from(container)
    .save()
    .then(() => container.remove());
}

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (character) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[character] ?? character);
}

function buildPreviewHtml(document: BusinessDocument, settings: BusinessSettings): string {
  const subtotal = subtotalOf(document);
  const iva = subtotal * 0.16;
  const total = subtotal + iva;
  const dt = new Date(`${document.date}T12:00:00`);
  const dateStr = dt.toLocaleDateString('es-MX', { day: 'numeric', month: 'short', year: '2-digit' }).replace('.', '');
  const folio = documentFolio(document);
  const logo = settings.logo || '/logo-ferreteria-garcia.jpg';
  const border = 'border:1px solid #000;';
  const header = `background:#00B0F0;color:#fff;font-weight:bold;text-align:center;${border}height:18px;`;
  const cell = `${border}height:16px;padding:0 3px;`;
  const itemRows: string[] = [];

  for (let i = 0; i < 13; i++) {
    const row = 15 + i;
    const item = document.items[i];
    itemRows.push(item
      ? `<tr style="height:16px;"><td style="${cell}text-align:center;">${i + 1}</td><td style="${cell}text-align:center;">${item.quantity}</td><td style="${cell}text-align:center;">${escapeHtml(item.unit)}</td><td colspan="2" style="${cell}">${escapeHtml(item.description)}</td><td style="${cell}text-align:right;white-space:nowrap;">${fmtMoney(item.unitPrice)}</td><td style="${cell}text-align:right;white-space:nowrap;">${fmtMoney(item.quantity * item.unitPrice)}</td></tr>`
      : `<tr style="height:16px;"><td style="${cell}"></td><td style="${cell}"></td><td style="${cell}"></td><td colspan="2" style="${cell}"></td><td style="${cell}"></td><td style="${cell}"></td></tr>`);
  }

  return `<div id="pdf-page" style="width:216mm;height:279mm;padding:10mm 12mm;box-sizing:border-box;font-family:Calibri,sans-serif;color:#000;font-size:10pt;background:#fff;">
  <table style="width:100%;table-layout:fixed;border-collapse:collapse;">
    <colgroup><col style="width:6.21%"><col style="width:10.77%"><col style="width:9.94%"><col style="width:35.20%"><col style="width:6.83%"><col style="width:13.46%"><col style="width:17.60%"></colgroup>
    <tr style="height:25px;"><td colspan="3" rowspan="7" style="vertical-align:middle;text-align:center;"><img src="${escapeHtml(logo)}" crossorigin="anonymous" style="max-width:100%;max-height:78px;object-fit:contain;" /></td><td></td><td colspan="3" rowspan="2" style="vertical-align:middle;text-align:right;font-family:'Century Gothic',sans-serif;font-size:22pt;font-weight:bold;color:#00B0F0;">${documentTitle(document.type)}</td></tr>
    <tr style="height:25px;"><td></td></tr>
    <tr style="height:25px;"><td></td><td style="font-weight:bold;text-align:right;padding-right:3px;">FECHA:</td><td colspan="2" style="${border}text-align:center;">${dateStr}</td></tr>
    <tr style="height:25px;"><td></td><td style="font-weight:bold;text-align:right;padding-right:3px;">FOLIO:</td><td colspan="2" style="${border}text-align:center;">#${folio}</td></tr>
    <tr style="height:25px;"><td colspan="4"></td></tr><tr style="height:25px;"><td colspan="4"></td></tr><tr style="height:9px;"><td colspan="4" style="border-bottom:1px solid #F00;"></td></tr>
  </table>
  <table style="width:100%;table-layout:fixed;border-collapse:collapse;font-size:9pt;">
    <colgroup><col style="width:6.21%"><col style="width:10.77%"><col style="width:9.94%"><col style="width:35.20%"><col style="width:6.83%"><col style="width:13.46%"><col style="width:17.60%"></colgroup>
    <tr style="height:16px;"><td colspan="2" style="text-align:right;font-weight:bold;">Cliente:</td><td colspan="2" style="font-size:10pt;">${escapeHtml(document.clientName)}</td><td colspan="3"></td></tr>
    <tr style="height:16px;"><td colspan="2" style="text-align:right;font-weight:bold;">Atención:</td><td colspan="2" style="font-size:10pt;">${escapeHtml(document.attention || '')}</td><td colspan="3"></td></tr>
    <tr style="height:16px;"><td colspan="2" style="text-align:right;font-weight:bold;">Correo Electrónico</td><td colspan="2" style="font-size:10pt;">${escapeHtml(document.email || '')}</td><td colspan="3"></td></tr>
    <tr style="height:16px;"><td colspan="2" style="text-align:right;font-weight:bold;">Teléfono:</td><td colspan="2" style="font-size:10pt;">${escapeHtml(document.phone || '')}</td><td colspan="3"></td></tr>
  </table>
  <table style="width:100%;table-layout:fixed;border-collapse:collapse;font-size:9pt;margin-top:3px;">
    <colgroup><col style="width:6.21%"><col style="width:10.77%"><col style="width:9.94%"><col style="width:35.20%"><col style="width:6.83%"><col style="width:13.46%"><col style="width:17.60%"></colgroup>
    <tr><td style="${header}">No.</td><td style="${header}">Cantidad</td><td style="${header}">Unidad</td><td colspan="2" style="${header}">Descripción</td><td style="${header}">Precio Unit.</td><td style="${header}">Total</td></tr>
    ${itemRows.join('')}
    <tr style="height:16px;"><td colspan="5"></td><td style="${cell}font-weight:bold;text-align:right;">SUBTOTAL</td><td style="${cell}text-align:right;">${fmtMoney(subtotal)}</td></tr>
    <tr style="height:16px;"><td colspan="5"></td><td style="${cell}font-weight:bold;text-align:right;">IVA</td><td style="${cell}text-align:right;">${fmtMoney(iva)}</td></tr>
    <tr style="height:16px;"><td colspan="5"></td><td style="${cell}font-weight:bold;text-align:right;">TOTAL</td><td style="${cell}text-align:right;">${fmtMoney(total)}</td></tr>
  </table>
  <table style="width:100%;table-layout:fixed;border-collapse:collapse;font-size:10pt;margin-top:0;">
    <colgroup><col style="width:6.21%"><col style="width:10.77%"><col style="width:9.94%"><col style="width:35.20%"><col style="width:6.83%"><col style="width:13.46%"><col style="width:17.60%"></colgroup>
    <tr style="height:18px;"><td colspan="4" style="background:#00B0F0;color:#fff;font-weight:bold;text-align:center;">Observaciones</td><td colspan="3"></td></tr>
    <tr style="height:19px;"><td colspan="3" style="text-align:right;font-weight:bold;">Tiempo de Entrega:</td><td colspan="2">${escapeHtml(document.delivery)}</td><td colspan="2"></td></tr>
    <tr style="height:19px;"><td colspan="3" style="text-align:right;font-weight:bold;">Forma de Pago:</td><td colspan="2">${escapeHtml(document.payment)}</td><td colspan="2"></td></tr>
    <tr style="height:19px;"><td colspan="3" style="text-align:right;font-weight:bold;">Nota:</td><td colspan="2" style="border-bottom:2px solid #000;">${escapeHtml(document.note || '')}</td><td colspan="2"></td></tr>
  </table>
  <table style="width:100%;table-layout:fixed;border-collapse:collapse;margin-top:42px;"><colgroup><col style="width:6.21%"><col style="width:10.77%"><col style="width:9.94%"><col style="width:35.20%"><col style="width:6.83%"><col style="width:13.46%"><col style="width:17.60%"></colgroup><tr><td colspan="3"></td><td colspan="2" style="border-top:1px solid #000;text-align:center;padding-top:4px;font-weight:bold;">Sello o firma de recibido</td><td colspan="2"></td></tr></table>
  <table style="width:100%;table-layout:fixed;border-collapse:collapse;margin-top:8px;"><tr><td style="text-align:center;font-size:12pt;font-weight:bold;">Ventas: ${escapeHtml(settings.salesName)}, ${escapeHtml(settings.salesPhone)}. Correo: ${escapeHtml(settings.salesEmail)}</td></tr></table>
</div>`;
}

function fmtMoney(value: number): string {
  return '$ ' + value.toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
