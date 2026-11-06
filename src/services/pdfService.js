const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer-core');
const QRCode = require('qrcode');
const { numberToWords } = require('./numberToWords');

// Detect available system Chromium/Edge/Chrome
function getBrowserExecutablePath() {
  const possiblePaths = [
    'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
    'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
    'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
    '/usr/bin/google-chrome',
    '/usr/bin/chromium-browser',
    '/usr/bin/chromium'
  ];

  for (const p of possiblePaths) {
    if (fs.existsSync(p)) {
      return p;
    }
  }
  return null;
}

// Format Currency in Indian Numbering Format ₹1,00,000.00
function formatINR(amount) {
  const num = Number(amount) || 0;
  const isNegative = num < 0;
  const absNum = Math.abs(num);
  const [intPart, decPart] = absNum.toFixed(2).split('.');

  let lastThree = intPart.slice(-3);
  let otherNumbers = intPart.slice(0, -3);
  if (otherNumbers !== '') {
    lastThree = ',' + lastThree;
  }
  const formattedInt = otherNumbers.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + lastThree;
  return `${isNegative ? '-' : ''}₹${formattedInt}.${decPart}`;
}

// Format Display Date DD/MM/YYYY
function formatDisplayDate(d) {
  if (!d) return '-';
  const str = String(d).trim();
  if (!str) return '-';
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(str)) return str;
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
    const p = str.slice(0, 10).split('-');
    return `${p[2]}/${p[1]}/${p[0]}`;
  }
  if (/^\d{2}-\d{2}-\d{4}$/.test(str)) {
    const p = str.split('-');
    return `${p[0]}/${p[1]}/${p[2]}`;
  }
  const dt = new Date(str);
  if (isNaN(dt.getTime())) return str;
  const dd = String(dt.getDate()).padStart(2, '0');
  const mm = String(dt.getMonth() + 1).padStart(2, '0');
  const yyyy = dt.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

// Convert Logo file to Base64
function getLogoBase64(logoPath) {
  try {
    const candidates = [
      path.resolve(__dirname, '..', '..', '..', 'Yogteck-Frontend', 'public', logoPath || 'assets/logo/yogteck-logo.webp'),
      path.resolve(__dirname, '..', '..', '..', 'Yogteck-Frontend', 'public', 'assets', 'logo', 'yogteck-logo.webp'),
      path.resolve(__dirname, '..', '..', 'data', 'logo.png')
    ];

    for (const file of candidates) {
      if (fs.existsSync(file)) {
        const ext = path.extname(file).replace('.', '').toLowerCase();
        const mime = ext === 'svg' ? 'image/svg+xml' : ext === 'webp' ? 'image/webp' : ext === 'png' ? 'image/png' : 'image/jpeg';
        const buffer = fs.readFileSync(file);
        return `data:${mime};base64,${buffer.toString('base64')}`;
      }
    }
  } catch (e) {
    console.warn('[PDF LOGO]', e.message);
  }
  return null;
}

// Convert Payment PhonePe QR file to Base64
function getPaymentQrBase64() {
  try {
    const candidates = [
      path.resolve(__dirname, '..', '..', 'data', 'payment-qr.jpg'),
      path.resolve(__dirname, '..', '..', '..', 'Yogteck-Frontend', 'public', 'assets', 'payment-qr.jpg')
    ];
    for (const f of candidates) {
      if (fs.existsSync(f)) {
        const buffer = fs.readFileSync(f);
        return `data:image/jpeg;base64,${buffer.toString('base64')}`;
      }
    }
  } catch (e) {
    console.warn('[PAYMENT QR]', e.message);
  }
  return null;
}

// Generate QR Code Data URL
async function generateQRCodeDataUrl(text) {
  try {
    return await QRCode.toDataURL(text, {
      margin: 1,
      width: 140,
      color: {
        dark: '#050B1A',
        light: '#FFFFFF'
      }
    });
  } catch (err) {
    console.error('QR generation error:', err);
    return '';
  }
}

/**
 * Generate Invoice HTML Template matching Gujarat Freight Tools sample layout
 * with Bilingual Important Notes Section at bottom
 */
async function buildInvoiceHtml(invoice, company, baseUrl = 'https://yogteck.com') {
  const isGstEnabled = String(company.gstEnabled || 'Yes').toLowerCase() === 'yes';
  const companyStateCode = String(company.stateCode || '09').trim();
  const clientStateCode = String(invoice.client?.stateCode || '').trim();
  const isInterState = companyStateCode !== clientStateCode && Boolean(clientStateCode);

  const logoBase64 = getLogoBase64(company.logoPath);
  const paymentQrBase64 = getPaymentQrBase64();

  // 1. Balance QR (links to public invoice view)
  const publicViewUrl = `${baseUrl}/invoice/view/${invoice.publicToken || 'sample-token'}`;
  const balanceQrUrl = await generateQRCodeDataUrl(publicViewUrl);

  // 2. UPI QR (optional upi://pay URI)
  let upiQrUrl = '';
  if (company.upiId && invoice.balanceDue > 0) {
    const upiUri = `upi://pay?pa=${encodeURIComponent(company.upiId)}&pn=${encodeURIComponent(company.companyName)}&am=${encodeURIComponent(invoice.balanceDue.toFixed(2))}&cu=INR&tn=${encodeURIComponent(invoice.invoiceNo)}`;
    upiQrUrl = await generateQRCodeDataUrl(upiUri);
  }

  // Words representation
  const grandTotalWords = numberToWords(invoice.grandTotal);

  // Calculate tax split
  const totalTax = Number(invoice.taxTotal) || 0;
  const taxableTotal = Number(invoice.taxableTotal) || 0;
  const halfTax = (totalTax / 2).toFixed(2);

  // Items rows
  const items = Array.isArray(invoice.items) ? invoice.items : [];
  let totalQty = 0;
  items.forEach(it => { totalQty += (Number(it.qty) || 0); });

  // Keep filler rows compact so that Important Notes fits cleanly
  const minRows = 3;
  const emptyRowsCount = Math.max(0, minRows - items.length);

  // Status color
  const statusColor = invoice.status === 'Paid' ? '#16a34a' : invoice.status === 'Partially Paid' ? '#d97706' : '#dc2626';
  const statusBg = invoice.status === 'Paid' ? '#dcfce7' : invoice.status === 'Partially Paid' ? '#fef3c7' : '#fee2e2';

  const termsLines = (company.terms || '').split('\n').filter(Boolean);

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>${invoice.invoiceNo}</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=Noto+Sans+Devanagari:wght@400;500;600;700;800&display=swap');

  @page {
    size: A4 portrait;
    margin: 8mm 9mm 8mm 9mm;
  }

  * {
    box-sizing: border-box;
    -webkit-print-color-adjust: exact !important;
    print-color-adjust: exact !important;
  }

  body {
    font-family: 'Inter', 'Noto Sans Devanagari', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    color: #111827;
    margin: 0;
    padding: 0;
    font-size: 11px;
    line-height: 1.35;
    background: #ffffff;
  }

  .invoice-container {
    width: 100%;
    max-width: 790px;
    margin: 0 auto;
    border: 1.5px solid #000000;
  }

  /* Header Section */
  .header-table {
    width: 100%;
    border-collapse: collapse;
  }

  .header-left {
    width: 62%;
    padding: 10px 14px 6px 14px;
    vertical-align: top;
  }

  .company-title {
    font-size: 23px;
    font-weight: 900;
    color: #0A1330;
    letter-spacing: 0.5px;
    text-transform: uppercase;
    margin: 0 0 5px 0;
    line-height: 1.1;
  }

  .tagline-ribbon {
    background: #008779;
    color: #ffffff;
    font-size: 10px;
    font-weight: 700;
    padding: 3px 8px;
    border-radius: 3px;
    margin-bottom: 5px;
    display: inline-block;
    letter-spacing: 0.3px;
  }

  .company-address {
    font-size: 10px;
    color: #334155;
    line-height: 1.4;
  }

  .header-right {
    width: 38%;
    padding: 10px 14px 6px 10px;
    text-align: right;
    vertical-align: top;
  }

  .company-logo {
    max-height: 58px;
    max-width: 140px;
    object-fit: contain;
    margin-bottom: 4px;
  }

  .contact-meta {
    font-size: 9.5px;
    color: #334155;
    line-height: 1.4;
  }
  .contact-meta strong {
    color: #0f172a;
  }

  /* Bar: PAN | TAX INVOICE | ORIGINAL FOR RECIPIENT */
  .title-bar {
    width: 100%;
    border-top: 1.5px solid #000000;
    border-bottom: 1.5px solid #000000;
    display: table;
    table-layout: fixed;
    background: #ffffff;
  }

  .title-bar-cell {
    display: table-cell;
    vertical-align: middle;
    padding: 4px 12px;
    font-size: 10.5px;
    font-weight: 700;
  }

  .title-bar-left {
    width: 35%;
    text-align: left;
    text-transform: uppercase;
  }

  .title-bar-center {
    width: 30%;
    text-align: center;
    font-size: 13.5px;
    font-weight: 800;
    letter-spacing: 1px;
  }

  .title-bar-right {
    width: 35%;
    text-align: right;
    font-size: 9px;
    font-weight: 600;
    text-transform: uppercase;
    color: #1e293b;
  }

  /* Two Column Detail Box */
  .detail-table {
    width: 100%;
    border-collapse: collapse;
  }

  .detail-left {
    width: 48%;
    vertical-align: top;
    padding: 6px 10px;
    border-right: 1.5px solid #000000;
  }

  .detail-right {
    width: 52%;
    vertical-align: top;
    padding: 6px 10px;
  }

  .box-header-title {
    text-align: center;
    font-weight: 700;
    font-size: 10.5px;
    border-bottom: 1px solid #000000;
    padding-bottom: 2px;
    margin-bottom: 4px;
    text-transform: capitalize;
  }

  .meta-row {
    display: table;
    width: 100%;
    margin-bottom: 2px;
    font-size: 10px;
  }

  .meta-label {
    display: table-cell;
    font-weight: 700;
    color: #0f172a;
    width: 75px;
    vertical-align: top;
  }

  .meta-value {
    display: table-cell;
    vertical-align: top;
    color: #1e293b;
  }

  .right-meta-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 10px;
  }

  .right-meta-table td {
    padding: 2px 4px;
    vertical-align: top;
  }

  .rm-label {
    font-weight: 600;
    width: 32%;
    color: #1e293b;
  }

  .rm-val {
    font-weight: 700;
    width: 28%;
    color: #0f172a;
  }

  /* Items Table */
  .items-table {
    width: 100%;
    border-collapse: collapse;
    border-top: 1.5px solid #000000;
    border-bottom: 1.5px solid #000000;
    font-size: 10px;
  }

  .items-table th {
    background: #f8fafc;
    border: 1px solid #000000;
    padding: 5px 4px;
    font-weight: 700;
    text-align: center;
    vertical-align: middle;
  }

  .items-table td {
    border-left: 1px solid #000000;
    border-right: 1px solid #000000;
    padding: 4px 6px;
    vertical-align: top;
  }

  .items-table .row-item {
    border-bottom: 1px solid #e2e8f0;
  }

  .items-table .empty-row td {
    height: 18px;
  }

  .items-table .total-row td {
    border-top: 1.5px solid #000000;
    border-bottom: 1.5px solid #000000;
    font-weight: 700;
    background: #f8fafc;
    padding: 5px 6px;
  }

  .text-center { text-align: center; }
  .text-right { text-align: right; }
  .text-left { text-align: left; }

  /* Bottom Split Section */
  .bottom-table {
    width: 100%;
    border-collapse: collapse;
  }

  .bottom-left-cell {
    width: 58%;
    vertical-align: top;
    border-right: 1.5px solid #000000;
    padding: 0;
  }

  .bottom-right-cell {
    width: 42%;
    vertical-align: top;
    padding: 0;
  }

  /* Total in words */
  .words-section {
    border-bottom: 1px solid #000000;
    padding: 5px 8px;
  }

  .words-title {
    text-align: center;
    font-weight: 700;
    font-size: 10px;
    margin-bottom: 2px;
    text-transform: capitalize;
  }

  .words-content {
    font-weight: 700;
    text-align: center;
    font-size: 9.5px;
    color: #0f172a;
    text-transform: uppercase;
  }

  /* Bank Details + QR Row */
  .bank-qr-table {
    width: 100%;
    border-collapse: collapse;
    border-bottom: 1px solid #000000;
  }

  .bank-details-cell {
    width: 65%;
    padding: 5px 8px;
    vertical-align: top;
  }

  .bank-title {
    text-align: center;
    font-weight: 700;
    font-size: 10px;
    border-bottom: 1px solid #000000;
    padding-bottom: 2px;
    margin-bottom: 4px;
  }

  .bank-row {
    font-size: 9.5px;
    margin-bottom: 2px;
    display: table;
    width: 100%;
  }

  .bank-lbl {
    display: table-cell;
    width: 70px;
    font-weight: 600;
    color: #475569;
  }

  .bank-val {
    display: table-cell;
    font-weight: 700;
    color: #0f172a;
  }

  .qr-cell {
    width: 35%;
    padding: 4px;
    text-align: center;
    vertical-align: middle;
    border-left: 1px solid #e2e8f0;
  }

  .qr-img {
    width: 66px;
    height: 66px;
    display: block;
    margin: 0 auto;
  }

  .qr-caption {
    font-size: 8px;
    font-weight: 700;
    color: #0A1330;
    margin-top: 2px;
  }

  /* Terms and Conditions */
  .terms-section {
    padding: 5px 8px;
    font-size: 9px;
    color: #334155;
    border-bottom: 1px solid #000000;
  }

  .terms-title {
    font-weight: 700;
    font-size: 9.5px;
    color: #0f172a;
    text-align: center;
    margin-bottom: 2px;
  }

  .customer-sig-box {
    padding: 16px 8px 4px 8px;
    font-size: 9.5px;
    font-weight: 600;
    color: #0f172a;
  }

  /* Right Side: Tax Summary + Payments Table */
  .calc-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 10px;
  }

  .calc-table td {
    padding: 3px 8px;
    border-bottom: 1px solid #000000;
  }

  .calc-label {
    font-weight: 600;
    color: #1e293b;
    text-align: left;
  }

  .calc-val {
    font-weight: 700;
    text-align: right;
    color: #0f172a;
  }

  .total-amt-row td {
    font-size: 12px;
    font-weight: 900;
    background: #f8fafc;
    border-bottom: 1.5px solid #000000;
    color: #0A1330;
  }

  .eoe-tag {
    font-size: 8.5px;
    font-weight: 600;
    text-align: right;
    padding: 2px 8px;
    color: #64748b;
    border-bottom: 1px solid #000000;
  }

  /* Payments Received Table */
  .payments-section {
    padding: 5px 8px;
    border-bottom: 1px solid #000000;
    background: #fafafa;
  }

  .payments-title {
    font-weight: 700;
    font-size: 9.5px;
    color: #0f172a;
    margin-bottom: 3px;
    display: flex;
    justify-content: space-between;
    align-items: center;
  }

  .status-badge {
    display: inline-block;
    padding: 1px 5px;
    border-radius: 4px;
    font-size: 8.5px;
    font-weight: 700;
    color: ${statusColor};
    background: ${statusBg};
    border: 1px solid ${statusColor}40;
  }

  .payments-mini-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 8.5px;
    margin-bottom: 3px;
  }

  .payments-mini-table th {
    background: #e2e8f0;
    border: 1px solid #cbd5e1;
    padding: 2px 4px;
    text-align: center;
  }

  .payments-mini-table td {
    border: 1px solid #cbd5e1;
    padding: 2px 4px;
  }

  .payment-summary-line {
    display: flex;
    justify-content: space-between;
    font-size: 9.5px;
    margin-top: 2px;
  }

  .due-highlight {
    font-size: 11px;
    font-weight: 800;
    color: ${invoice.balanceDue > 0 ? '#b91c1c' : '#15803d'};
  }

  /* Signatory Block */
  .signatory-box {
    padding: 5px 8px 4px 8px;
    text-align: center;
  }

  .cert-text {
    font-size: 7.5px;
    color: #475569;
    margin-bottom: 2px;
  }

  .for-company {
    font-size: 10.5px;
    font-weight: 800;
    color: #0A1330;
    margin-bottom: 16px;
  }

  .slanted-note {
    font-size: 8px;
    color: #475569;
    font-style: italic;
    transform: rotate(-3deg);
    margin: 4px 0 10px 0;
  }

  .auth-sign-label {
    font-size: 9px;
    font-weight: 700;
    border-top: 1px solid #94a3b8;
    padding-top: 2px;
    display: inline-block;
    width: 70%;
  }

  /* ------------------------------------------------------------- */
  /* SECTION 9: IMPORTANT NOTES BOX (BILINGUAL WITH NOTO SANS)     */
  /* ------------------------------------------------------------- */
  .important-notes-box {
    border-top: 1.5px solid #000000;
    background: #ffffff;
    page-break-inside: avoid;
    break-inside: avoid;
  }

  .notes-hdr-bar {
    background: #0A1330;
    color: #ffffff;
    padding: 4px 10px;
    font-size: 10px;
    font-weight: 800;
    letter-spacing: 0.5px;
    text-transform: uppercase;
  }

  .notes-content {
    padding: 6px 10px;
  }

  .note-item {
    margin-bottom: 6px;
    padding-bottom: 4px;
    border-bottom: 1px dashed #cbd5e1;
  }

  .note-item:last-child {
    margin-bottom: 0;
    padding-bottom: 0;
    border-bottom: none;
  }

  .note-title {
    font-weight: 800;
    font-size: 9.5px;
    color: #0f172a;
    margin-bottom: 2px;
  }

  .note-hi {
    font-family: 'Noto Sans Devanagari', sans-serif;
    color: #1e293b;
    font-size: 9px;
    line-height: 1.4;
    margin: 0 0 2px 0;
  }

  .note-en {
    color: #475569;
    font-style: italic;
    font-size: 8.5px;
    line-height: 1.35;
    margin: 0;
  }

  /* Bottom Thank You Banner */
  .thank-you-footer {
    padding: 4px;
    text-align: center;
    font-size: 9.5px;
    font-weight: 700;
    color: #334155;
    border-top: 1px solid #000000;
    background: #ffffff;
    letter-spacing: 0.5px;
  }
</style>
</head>
<body>

<div class="invoice-container">
  
  <!-- 1. Header Section -->
  <table class="header-table">
    <tr>
      <td class="header-left">
        <div class="company-title">${company.companyName}</div>
        <div class="tagline-ribbon">${company.tagline || 'Complete Digital Solutions'}</div>
        <div class="company-address">
          ${company.address}<br>
          Pincode : ${company.pincode} &bull; State : ${company.state} (${company.stateCode})
        </div>
      </td>
      <td class="header-right">
        ${logoBase64 ? `<img src="${logoBase64}" class="company-logo" alt="YOGTECK Logo">` : ''}
        <div class="contact-meta">
          <div><strong>Tel :</strong> ${company.contact}</div>
          <div><strong>Web :</strong> www.yogteck.com</div>
          <div><strong>Email :</strong> ${company.email}</div>
        </div>
      </td>
    </tr>
  </table>

  <!-- 2. Border Strip: PAN | TAX INVOICE | ORIGINAL FOR RECIPIENT -->
  <div class="title-bar">
    <div class="title-bar-cell title-bar-left">
      PAN : <strong>${company.pan}</strong>
    </div>
    <div class="title-bar-cell title-bar-center">
      ${isGstEnabled ? 'TAX INVOICE' : 'INVOICE'}
    </div>
    <div class="title-bar-cell title-bar-right">
      ORIGINAL FOR RECIPIENT
    </div>
  </div>

  <!-- 3. Two Column Details Box -->
  <table class="detail-table">
    <tr>
      <!-- Customer Detail -->
      <td class="detail-left">
        <div class="box-header-title">Customer Detail</div>
        <div class="meta-row">
          <div class="meta-label">M/S</div>
          <div class="meta-value"><strong>${invoice.client?.clientName || 'N/A'}</strong></div>
        </div>
        <div class="meta-row">
          <div class="meta-label">Address</div>
          <div class="meta-value">${invoice.client?.address || 'Not Specified'}</div>
        </div>
        <div class="meta-row">
          <div class="meta-label">Phone</div>
          <div class="meta-value">${invoice.client?.mobile || 'Not Specified'}</div>
        </div>
        ${isGstEnabled ? `
        <div class="meta-row">
          <div class="meta-label">GSTIN</div>
          <div class="meta-value">${invoice.client?.gstin || 'URP (Unregistered)'}</div>
        </div>
        <div class="meta-row">
          <div class="meta-label">Place of Supply</div>
          <div class="meta-value">${invoice.client?.state || company.state} ( ${invoice.client?.stateCode || company.stateCode} )</div>
        </div>` : ''}
      </td>

      <!-- Invoice Metadata Detail -->
      <td class="detail-right">
        <table class="right-meta-table">
          <tr>
            <td class="rm-label">Invoice No.</td>
            <td class="rm-val">${invoice.invoiceNo}</td>
            <td class="rm-label">Invoice Date</td>
            <td class="rm-val">${formatDisplayDate(invoice.invoiceDate)}</td>
          </tr>
          ${invoice.planName ? `
          <tr>
            <td class="rm-label">Plan Name</td>
            <td class="rm-val" colspan="3"><strong>${invoice.planName}</strong></td>
          </tr>` : ''}
          ${invoice.serviceStartDate ? `
          <tr>
            <td class="rm-label">Plan Active</td>
            <td class="rm-val">${formatDisplayDate(invoice.serviceStartDate)}</td>
            <td class="rm-label">Plan Expiry</td>
            <td class="rm-val">${formatDisplayDate(invoice.serviceExpiryDate)}</td>
          </tr>` : ''}
          ${invoice.challanNo ? `
          <tr>
            <td class="rm-label">Challan No</td>
            <td class="rm-val">${invoice.challanNo}</td>
            <td class="rm-label">Challan Date</td>
            <td class="rm-val">${formatDisplayDate(invoice.challanDate)}</td>
          </tr>` : ''}
          ${invoice.ewayBillNo ? `
          <tr>
            <td class="rm-label">E-Way Bill No.</td>
            <td class="rm-val" colspan="3">${invoice.ewayBillNo}</td>
          </tr>` : ''}
          ${invoice.transport ? `
          <tr>
            <td class="rm-label">Transport</td>
            <td class="rm-val">${invoice.transport}</td>
            <td class="rm-label">Transport ID</td>
            <td class="rm-val">${invoice.transportId || '-'}</td>
          </tr>` : ''}
          <tr>
            <td class="rm-label">Due Date</td>
            <td class="rm-val" colspan="3">${formatDisplayDate(invoice.dueDate || invoice.renewalDueDate)}</td>
          </tr>
        </table>
      </td>
    </tr>
  </table>

  <!-- 4. Items Table (No HSN, No Tax; With Payment Date) -->
  <table class="items-table">
    <thead>
      <tr>
        <th style="width: 6%;">Sr.<br>No.</th>
        <th style="width: 44%;">Name of Product / Service</th>
        <th style="width: 16%;">Payment Date</th>
        <th style="width: 10%;">Qty</th>
        <th style="width: 12%;">Rate</th>
        <th style="width: 12%;">Total</th>
      </tr>
    </thead>
    <tbody>
      ${items.map((item, idx) => `
        <tr class="row-item">
          <td class="text-center">${idx + 1}</td>
          <td class="text-left"><strong>${item.description}</strong></td>
          <td class="text-center">${formatDisplayDate(item.paymentDate || invoice.paymentDate || invoice.invoiceDate)}</td>
          <td class="text-center">${item.qty} ${item.unit || 'NOS'}</td>
          <td class="text-right">${formatINR(item.rate)}</td>
          <td class="text-right"><strong>${formatINR(item.lineTotal || (item.qty * item.rate))}</strong></td>
        </tr>
      `).join('')}

      <!-- Empty filler rows to maintain table height -->
      ${Array.from({ length: emptyRowsCount }).map(() => `
        <tr class="empty-row">
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
      `).join('')}

      <!-- Total Row -->
      <tr class="total-row">
        <td colspan="4" class="text-right" style="font-weight: 800;">Total Amount</td>
        <td colspan="2" class="text-right"><strong>${formatINR(invoice.grandTotal)}</strong></td>
      </tr>
    </tbody>
  </table>

  <!-- 5. Bottom Split Section -->
  <table class="bottom-table">
    <tr>
      <!-- Left Cell -->
      <td class="bottom-left-cell">
        <!-- Total in words -->
        <div class="words-section">
          <div class="words-title">Total in words</div>
          <div class="words-content">${grandTotalWords}</div>
        </div>

        <!-- Bank Details + QR Code Table -->
        <table class="bank-qr-table">
          <tr>
            <td class="bank-details-cell">
              <div class="bank-title">Bank Details</div>
              <div class="bank-row"><span class="bank-lbl">Name</span> <span class="bank-val">${company.bankName || 'BANK OF INDIA'}</span></div>
              <div class="bank-row"><span class="bank-lbl">Branch</span> <span class="bank-val">${company.branch || 'Kanpur'}</span></div>
              <div class="bank-row"><span class="bank-lbl">Acc. Number</span> <span class="bank-val">${company.accountNo || '731410110007057'}</span></div>
              <div class="bank-row"><span class="bank-lbl">IFSC</span> <span class="bank-val">${company.ifsc || 'BKID0007336'}</span></div>
              ${company.upiId ? `<div class="bank-row"><span class="bank-lbl">UPI ID</span> <span class="bank-val">${company.upiId}</span></div>` : ''}
            </td>
            <td class="qr-cell">
              ${paymentQrBase64 ? `<img src="${paymentQrBase64}" class="qr-img" style="width: 78px; height: 78px; object-fit: contain;" alt="UPI QR">` : (balanceQrUrl ? `<img src="${balanceQrUrl}" class="qr-img" alt="QR Code">` : '')}
              <div class="qr-caption" style="font-weight: 800; color: #0A1330; text-transform: uppercase;">Pay using UPI</div>
            </td>
          </tr>
        </table>

        <!-- Terms and Conditions -->
        <div class="terms-section">
          <div class="terms-title">Terms and Conditions</div>
          ${termsLines.map(line => `<div>${line}</div>`).join('')}
        </div>

        <!-- Customer Signature -->
        <div class="customer-sig-box">
          Customer Signature
        </div>
      </td>

      <!-- Right Cell: Calculations, Payments Received, Signatory -->
      <td class="bottom-right-cell">
        <table class="calc-table">
          <tr class="total-amt-row">
            <td class="calc-label" style="font-size: 11px;">Total Amount</td>
            <td class="calc-val" style="font-size: 12.5px;">${formatINR(invoice.grandTotal)}</td>
          </tr>
        </table>

        <div class="eoe-tag">(E &amp; O.E.)</div>

        <!-- Payments Received Table -->
        <div class="payments-section">
          <div class="payments-title">
            <span>Payment Details</span>
            <span class="status-badge">${invoice.status}</span>
          </div>

          ${Array.isArray(invoice.payments) && invoice.payments.length > 0 ? `
            <table class="payments-mini-table">
              <thead>
                <tr>
                  <th>Sr</th>
                  <th>Date</th>
                  <th>Mode</th>
                  <th>Amount</th>
                </tr>
              </thead>
              <tbody>
                ${invoice.payments.map((p, pIdx) => `
                  <tr>
                    <td class="text-center">${pIdx + 1}</td>
                    <td class="text-center">${formatDisplayDate(p.paymentDate)}</td>
                    <td class="text-center">${p.mode || 'UPI'}</td>
                    <td class="text-right"><strong>${formatINR(p.amountPaid)}</strong></td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          ` : `
            <div style="font-size: 8.5px; color: #64748b; font-style: italic; margin-bottom: 3px;">
              ${invoice.status === 'Paid' ? `Full Payment Recorded on ${formatDisplayDate(invoice.paymentDate || invoice.invoiceDate)}.` : 'No payments recorded yet.'}
            </div>
          `}
            <table class="payments-mini-table">
              <thead>
                <tr>
                  <th>Sr</th>
                  <th>Date</th>
                  <th>Mode</th>
                  <th>Amount</th>
                </tr>
              </thead>
              <tbody>
                ${invoice.payments.map((p, pIdx) => `
                  <tr>
                    <td class="text-center">${pIdx + 1}</td>
                    <td class="text-center">${p.paymentDate}</td>
                    <td class="text-center">${p.mode || 'UPI'}</td>
                    <td class="text-right"><strong>${formatINR(p.amountPaid)}</strong></td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          ` : `
            <div style="font-size: 8.5px; color: #64748b; font-style: italic; margin-bottom: 3px;">No payments recorded yet.</div>
          `}

          <div class="payment-summary-line">
            <span style="color:#475569; font-weight:600;">Total Paid:</span>
            <span style="font-weight:700;">${formatINR(invoice.totalPaid || 0)}</span>
          </div>
          <div class="payment-summary-line" style="margin-top: 2px;">
            <span style="font-weight:700;">Balance Due:</span>
            <span class="due-highlight">${formatINR(invoice.balanceDue != null ? invoice.balanceDue : invoice.grandTotal)}</span>
          </div>
        </div>

        <!-- Signatory Block -->
        <div class="signatory-box">
          <div class="cert-text">Certified that the particulars given above are true and correct.</div>
          <div class="for-company">For ${company.companyName}</div>
          <div class="slanted-note">This is a computer generated invoice, no signature required.</div>
          <div class="auth-sign-label">Authorised Signatory</div>
        </div>

      </td>
    </tr>
  </table>

  <!-- SECTION 9: IMPORTANT NOTES / ज़रूरी सूचना (BILINGUAL) -->
  <div class="important-notes-box">
    <div class="notes-hdr-bar">
      ज़रूरी सूचना / Important Notes
    </div>
    <div class="notes-content">
      
      <!-- 1) भुगतान और सर्विस रिन्यूअल -->
      <div class="note-item">
        <div class="note-title">1) भुगतान और सर्विस रिन्यूअल / Payment &amp; Service Renewal</div>
        <div class="note-hi">
          यदि भुगतान या रिन्यूअल में देरी होती है, तो सर्विसेज़ की अवधि समाप्त (Expire) हो जाने के कारण आपका प्रोजेक्ट/वेबसाइट कभी भी बंद हो सकता है। इसलिए कृपया सर्विसेज़ को निर्बाध रूप से चालू रखने के लिए समय पर भुगतान या रिन्यूअल कर लें। देरी की स्थिति में किसी भी प्रकार के नुकसान या डेटा/सेवा बाधा के लिए कंपनी उत्तरदायी नहीं होगी।
        </div>
        <div class="note-en">
          Delay in payment or renewal may cause your services to expire and your project/website to go offline at any time. Please make timely payments or renewals to keep services active. The company will not be held responsible for any loss or service interruption caused by delayed payment.
        </div>
      </div>

      <!-- 2) सब्सक्रिप्शन की जानकारी -->
      <div class="note-item">
        <div class="note-title">2) सब्सक्रिप्शन की जानकारी / Subscription &amp; Plan Details</div>
        <div class="note-hi">
          अपने सब्सक्रिप्शन से जुड़ी पूरी जानकारी देखने के लिए इस इनवॉइस पर दिए गए QR कोड को स्कैन करें। वहाँ आपको पूरा इनवॉइस दिखेगा, जिसमें आपके सभी प्लान, उनकी शुरुआत और समाप्ति की तिथियाँ (Expiry), भुगतान का इतिहास और बाकी रकम (Balance Due) सहित सभी संबंधित विवरण उपलब्ध होंगे।
        </div>
        <div class="note-en">
          Scan the QR code on this invoice to view your full live invoice, active plan details, start and expiry dates, payment history, and balance due.
        </div>
      </div>

      <!-- 3) क्लाइंट द्वारा उपलब्ध कराई जाने वाली सामग्री -->
      <div class="note-item">
        <div class="note-title">3) क्लाइंट द्वारा उपलब्ध कराई जाने वाली सामग्री / Materials &amp; External Resources to be Provided by Client</div>
        <div class="note-hi">
          प्रोजेक्ट से जुड़ी सभी बाहरी (External) फाइलें, सेवाएँ और संसाधन क्लाइंट/पार्टी को ही उपलब्ध कराने होंगे। उदाहरण के लिए:<br>
          &bull; पेमेंट गेटवे (Payment Gateway) का अकाउंट/क्रेडेंशियल्स क्लाइंट द्वारा उपलब्ध कराया जाएगा।<br>
          &bull; प्रोजेक्ट से जुड़ी क्लाइंट की अपनी फाइलें, जैसे CAD ड्रॉइंग्स, लोगो, कंटेंट, फोटो और अन्य दस्तावेज़ (जैसे MMR Constructions के प्रोजेक्ट में उनकी CAD फाइलें) क्लाइंट को ही देनी होंगी।<br>
          इन सामग्रियों के अभाव या देरी से होने वाली किसी भी बाधा या विलंब की ज़िम्मेदारी कंपनी की नहीं होगी।
        </div>
        <div class="note-en">
          All external files, accounts, and resources required for the project (e.g. payment gateway credentials, CAD files, logo, content, photos) must be provided by the client; the company cannot be held responsible for delays or disruptions arising from their absence.
        </div>
      </div>

    </div>
  </div>

  <!-- Bottom Thank You Banner -->
  <div class="thank-you-footer">
    Thank you for your business!
  </div>

</div>

</body>
</html>`;
}

/**
 * Generate A4 PDF Buffer using Headless Chromium/Edge
 */
async function generateInvoicePdfBuffer(invoice, company, baseUrl) {
  const browserPath = getBrowserExecutablePath();
  if (!browserPath) {
    throw new Error('No compatible Chrome or Edge executable found on the system for PDF generation.');
  }

  const html = await buildInvoiceHtml(invoice, company, baseUrl);

  const browser = await puppeteer.launch({
    executablePath: browserPath,
    headless: true,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-gpu',
      '--font-render-hinting=none'
    ]
  });

  try {
    const page = await browser.newPage();
    await page.setContent(html, {
      waitUntil: ['load', 'networkidle0'],
      timeout: 30000
    });

    const pdfBuffer = await page.pdf({
      format: 'A4',
      printBackground: true,
      margin: {
        top: '6mm',
        right: '8mm',
        bottom: '6mm',
        left: '8mm'
      }
    });

    return pdfBuffer;
  } finally {
    await browser.close().catch(() => {});
  }
}

module.exports = {
  buildInvoiceHtml,
  generateInvoicePdfBuffer,
  formatINR
};
