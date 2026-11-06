const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const ExcelJS = require('exceljs');

const defaultDataDir = process.env.VERCEL
  ? path.join('/tmp', 'yogteck-data')
  : path.join(__dirname, '..', '..', 'data');

const seedDataDir = path.join(__dirname, '..', '..', 'data');
const dataDir = process.env.DATA_DIR || defaultDataDir;
const CLIENTS_MASTER_PATH = path.join(dataDir, 'clients_master.xlsx');
const SEED_CLIENTS_FILE = path.join(seedDataDir, 'clients_master.xlsx');
const BACKUP_DIR = path.join(dataDir, 'backup');

// In-process Write Lock Queue
let writeQueue = Promise.resolve();

function executeWithLock(fn) {
  const result = writeQueue.then(fn, fn);
  writeQueue = result.catch(() => {});
  return result;
}

// Formula Injection Sanitization
function sanitize(val) {
  if (typeof val !== 'string') return val;
  const trimmed = val.trim();
  if (/^[=+\-@]/.test(trimmed)) {
    return `'${trimmed}`;
  }
  return trimmed;
}

function desanitize(val) {
  if (val == null) return '';
  if (typeof val === 'object' && val.text) {
    val = val.text;
  }
  const str = String(val).trim();
  if (str.startsWith("'") && /^[=+\-@]/.test(str.slice(1))) {
    return str.slice(1);
  }
  return str;
}

// Paise Precision Arithmetic Helpers
function toPaise(val) {
  return Math.round((Number(val) || 0) * 100);
}

function fromPaise(paise) {
  return (paise / 100).toFixed(2);
}

function roundToTwo(val) {
  return Number((Math.round((Number(val) || 0) * 100) / 100).toFixed(2));
}

// Backup Helper
async function backupBeforeWrite() {
  if (!fs.existsSync(CLIENTS_MASTER_PATH)) return;

  if (!fs.existsSync(BACKUP_DIR)) {
    try {
      fs.mkdirSync(BACKUP_DIR, { recursive: true });
    } catch (e) {
      console.warn('[BACKUP MKDIR]', e.message);
    }
  }

  const now = new Date();
  const timestamp = now.toISOString().replace(/[-:T.]/g, '').slice(0, 14);
  const backupPath = path.join(BACKUP_DIR, `clients_master_${timestamp}.xlsx`);

  try {
    fs.copyFileSync(CLIENTS_MASTER_PATH, backupPath);
  } catch (err) {
    console.warn('[BACKUP FAILED]', err.message);
  }
}

// Schema Columns Definitions
const SCHEMAS = {
  Clients: [
    { header: 'Client ID', key: 'clientId', width: 16 },
    { header: 'Client Name', key: 'clientName', width: 28 },
    { header: 'Owner/Contact Person', key: 'contactPerson', width: 24 },
    { header: 'Mobile', key: 'mobile', width: 16 },
    { header: 'Email', key: 'email', width: 28 },
    { header: 'Address', key: 'address', width: 45 },
    { header: 'State', key: 'state', width: 18 },
    { header: 'State Code', key: 'stateCode', width: 12 },
    { header: 'GSTIN', key: 'gstin', width: 18 },
    { header: 'Created Date', key: 'createdDate', width: 20 }
  ],
  Invoices: [
    { header: 'Invoice No', key: 'invoiceNo', width: 22 },
    { header: 'Client ID', key: 'clientId', width: 16 },
    { header: 'Invoice Date', key: 'invoiceDate', width: 14 },
    { header: 'Due Date', key: 'dueDate', width: 14 },
    { header: 'Taxable Total', key: 'taxableTotal', width: 16 },
    { header: 'Tax Total', key: 'taxTotal', width: 16 },
    { header: 'Grand Total', key: 'grandTotal', width: 16 },
    { header: 'Status', key: 'status', width: 16 },
    { header: 'Public Token', key: 'publicToken', width: 34 },
    { header: 'Notes', key: 'notes', width: 30 },
    { header: 'Challan No', key: 'challanNo', width: 16 },
    { header: 'Challan Date', key: 'challanDate', width: 14 },
    { header: 'EWay Bill No', key: 'ewayBillNo', width: 18 },
    { header: 'Transport', key: 'transport', width: 20 },
    { header: 'Transport ID', key: 'transportId', width: 20 },
    { header: 'Plan Name', key: 'planName', width: 28 },
    { header: 'Service Start Date', key: 'serviceStartDate', width: 16 },
    { header: 'Service Expiry Date', key: 'serviceExpiryDate', width: 16 },
    { header: 'Renewal Due Date', key: 'renewalDueDate', width: 16 }
  ],
  InvoiceItems: [
    { header: 'Invoice No', key: 'invoiceNo', width: 22 },
    { header: 'Sr No', key: 'srNo', width: 10 },
    { header: 'Description', key: 'description', width: 40 },
    { header: 'HSN/SAC', key: 'hsnSac', width: 14 },
    { header: 'Qty', key: 'qty', width: 10 },
    { header: 'Unit', key: 'unit', width: 10 },
    { header: 'Rate', key: 'rate', width: 14 },
    { header: 'Taxable Value', key: 'taxableValue', width: 16 },
    { header: 'Tax %', key: 'taxPercent', width: 10 },
    { header: 'Tax Amount', key: 'taxAmount', width: 14 },
    { header: 'Line Total', key: 'lineTotal', width: 16 }
  ],
  Payments: [
    { header: 'Payment ID', key: 'paymentId', width: 22 },
    { header: 'Invoice No', key: 'invoiceNo', width: 22 },
    { header: 'Amount Paid', key: 'amountPaid', width: 16 },
    { header: 'Payment Date', key: 'paymentDate', width: 14 },
    { header: 'Mode', key: 'mode', width: 16 },
    { header: 'Note', key: 'note', width: 30 }
  ]
};

function formatHeader(sheet) {
  const row = sheet.getRow(1);
  row.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  row.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF0A1330' }
  };
  row.alignment = { vertical: 'middle', horizontal: 'center' };
}

// Convert object to row values array matching schema
function objectToRowArray(sheetName, obj) {
  return SCHEMAS[sheetName].map(col => {
    const val = obj[col.key];
    return val != null ? val : '';
  });
}

async function loadWorkbook() {
  const workbook = new ExcelJS.Workbook();
  if (fs.existsSync(CLIENTS_MASTER_PATH)) {
    await workbook.xlsx.readFile(CLIENTS_MASTER_PATH);
  }
  for (const [sheetName, columns] of Object.entries(SCHEMAS)) {
    const sheet = workbook.getWorksheet(sheetName);
    if (sheet) {
      sheet.columns = columns;
    }
  }
  return workbook;
}

async function ensureClientsMasterInitialized() {
  if (!fs.existsSync(dataDir)) {
    try {
      fs.mkdirSync(dataDir, { recursive: true });
    } catch (e) {
      console.warn('[CLIENTS DIR MKDIR]', e.message);
    }
  }

  // If in Vercel/tmp and file not in /tmp yet, copy from seed if exists
  if (!fs.existsSync(CLIENTS_MASTER_PATH) && fs.existsSync(SEED_CLIENTS_FILE)) {
    try {
      fs.copyFileSync(SEED_CLIENTS_FILE, CLIENTS_MASTER_PATH);
      return;
    } catch (e) {
      console.warn('Could not copy seed clients file:', e.message);
    }
  }

  if (fs.existsSync(CLIENTS_MASTER_PATH)) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(CLIENTS_MASTER_PATH);
    let modified = false;

    for (const [sheetName, columns] of Object.entries(SCHEMAS)) {
      let sheet = workbook.getWorksheet(sheetName);
      if (!sheet) {
        sheet = workbook.addWorksheet(sheetName);
        sheet.columns = columns;
        formatHeader(sheet);
        modified = true;
      } else {
        const existingHeaders = [];
        sheet.getRow(1).eachCell((cell) => {
          existingHeaders.push(String(cell.value || '').trim().toLowerCase());
        });

        columns.forEach((col, idx) => {
          if (!existingHeaders.includes(col.header.toLowerCase())) {
            const cell = sheet.getRow(1).getCell(idx + 1);
            cell.value = col.header;
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FF0A1330' }
            };
            modified = true;
          }
        });
      }
    }

    if (modified) {
      await workbook.xlsx.writeFile(CLIENTS_MASTER_PATH);
    }
    return;
  }

  // Create new workbook with all 4 sheets
  const workbook = new ExcelJS.Workbook();
  for (const [sheetName, columns] of Object.entries(SCHEMAS)) {
    const sheet = workbook.addWorksheet(sheetName);
    sheet.columns = columns;
    formatHeader(sheet);
  }

  // Seed Initial Sample Client if newly created
  const clientsSheet = workbook.getWorksheet('Clients');
  const sampleClient = {
    clientId: 'YT-CLI-001',
    clientName: sanitize('MMR Constructions'),
    contactPerson: sanitize('Rajesh Sharma'),
    mobile: sanitize('9878789878'),
    email: sanitize('contact@mmrconstructions.in'),
    address: sanitize('Sumel Business Park 7, Ring Road, Ahmedabad, Gujarat - 380023'),
    state: sanitize('Gujarat'),
    stateCode: sanitize('24'),
    gstin: sanitize('24AABBA7890B1ZB'),
    createdDate: new Date().toISOString().slice(0, 10)
  };
  clientsSheet.addRow(objectToRowArray('Clients', sampleClient));

  await workbook.xlsx.writeFile(CLIENTS_MASTER_PATH);
}

// ----------------------------------------------------------------------
// READ OPERATIONS
// ----------------------------------------------------------------------

async function readSheetRows(sheetName) {
  await ensureClientsMasterInitialized();
  if (!fs.existsSync(CLIENTS_MASTER_PATH)) return [];

  const workbook = await loadWorkbook();
  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) return [];

  const columns = SCHEMAS[sheetName];
  const rows = [];

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const item = {};
    columns.forEach((col, idx) => {
      const cellVal = row.getCell(idx + 1).value;
      item[col.key] = desanitize(cellVal);
    });
    if (item[columns[0].key]) {
      rows.push(item);
    }
  });

  return rows;
}

async function getAllClients() {
  const rows = await readSheetRows('Clients');
  return rows.map(c => ({
    ...c,
    clientId: c.clientId,
    clientName: c.clientName,
    contactPerson: c.contactPerson,
    mobile: c.mobile,
    email: c.email,
    address: c.address,
    state: c.state,
    stateCode: c.stateCode,
    gstin: c.gstin,
    createdDate: c.createdDate
  }));
}

async function getClientById(clientId) {
  const clients = await getAllClients();
  return clients.find(c => c.clientId === clientId) || null;
}

async function getAllPayments() {
  const rows = await readSheetRows('Payments');
  return rows.map(p => ({
    ...p,
    amountPaid: roundToTwo(p.amountPaid)
  }));
}

async function getPaymentsForInvoice(invoiceNo) {
  const allPayments = await getAllPayments();
  return allPayments.filter(p => p.invoiceNo === invoiceNo);
}

async function getAllInvoiceItems() {
  const rows = await readSheetRows('InvoiceItems');
  return rows.map(item => ({
    ...item,
    srNo: Number(item.srNo) || 1,
    qty: Number(item.qty) || 1,
    rate: roundToTwo(item.rate),
    taxableValue: roundToTwo(item.taxableValue),
    taxPercent: Number(item.taxPercent) || 0,
    taxAmount: roundToTwo(item.taxAmount),
    lineTotal: roundToTwo(item.lineTotal)
  }));
}

async function getItemsForInvoice(invoiceNo) {
  const allItems = await getAllInvoiceItems();
  return allItems.filter(item => item.invoiceNo === invoiceNo);
}

async function getAllInvoices() {
  const [invoices, clients, payments] = await Promise.all([
    readSheetRows('Invoices'),
    getAllClients(),
    getAllPayments()
  ]);

  const clientMap = new Map(clients.map(c => [c.clientId, c]));

  const paymentsByInvoice = new Map();
  payments.forEach(p => {
    const list = paymentsByInvoice.get(p.invoiceNo) || [];
    list.push(p);
    paymentsByInvoice.set(p.invoiceNo, list);
  });

  return invoices.map(inv => {
    const invPayments = paymentsByInvoice.get(inv.invoiceNo) || [];
    const grandTotalPaise = toPaise(inv.grandTotal);
    const totalPaidPaise = invPayments.reduce((sum, p) => sum + toPaise(p.amountPaid), 0);
    const balanceDuePaise = Math.max(0, grandTotalPaise - totalPaidPaise);

    let status = 'Unpaid';
    if (totalPaidPaise >= grandTotalPaise && grandTotalPaise > 0) {
      status = 'Paid';
    } else if (totalPaidPaise > 0) {
      status = 'Partially Paid';
    }

    const client = clientMap.get(inv.clientId) || null;

    return {
      ...inv,
      taxableTotal: roundToTwo(inv.taxableTotal),
      taxTotal: roundToTwo(inv.taxTotal),
      grandTotal: roundToTwo(inv.grandTotal),
      totalPaid: roundToTwo(fromPaise(totalPaidPaise)),
      balanceDue: roundToTwo(fromPaise(balanceDuePaise)),
      status,
      client,
      paymentsCount: invPayments.length,
      latestPayment: invPayments[invPayments.length - 1] || null
    };
  });
}

async function getInvoiceByNo(invoiceNo) {
  const invoices = await getAllInvoices();
  const invoice = invoices.find(i => i.invoiceNo === invoiceNo);
  if (!invoice) return null;

  const [items, payments] = await Promise.all([
    getItemsForInvoice(invoiceNo),
    getPaymentsForInvoice(invoiceNo)
  ]);

  return {
    ...invoice,
    items,
    payments
  };
}

async function getInvoiceByToken(publicToken) {
  const invoices = await getAllInvoices();
  const invoice = invoices.find(i => i.publicToken === publicToken);
  if (!invoice) return null;

  const [items, payments] = await Promise.all([
    getItemsForInvoice(invoice.invoiceNo),
    getPaymentsForInvoice(invoice.invoiceNo)
  ]);

  return {
    ...invoice,
    items,
    payments
  };
}

// ----------------------------------------------------------------------
// WRITE OPERATIONS (Protected with Mutex Queue & Backup)
// ----------------------------------------------------------------------

async function generateNextInvoiceNo() {
  await ensureClientsMasterInitialized();
  const invoices = await readSheetRows('Invoices');

  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const prefix = `YOGTECK-${year}-${month}-`;

  let maxNum = 0;
  invoices.forEach(inv => {
    if (inv.invoiceNo && inv.invoiceNo.startsWith(prefix)) {
      const part = inv.invoiceNo.replace(prefix, '');
      const num = parseInt(part, 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  });

  const nextNum = String(maxNum + 1).padStart(3, '0');
  return `${prefix}${nextNum}`;
}

async function generateNextClientId() {
  const clients = await getAllClients();
  let maxNum = 0;
  clients.forEach(c => {
    if (c.clientId && c.clientId.startsWith('YT-CLI-')) {
      const num = parseInt(c.clientId.replace('YT-CLI-', ''), 10);
      if (!isNaN(num) && num > maxNum) maxNum = num;
    }
  });
  return `YT-CLI-${String(maxNum + 1).padStart(3, '0')}`;
}

async function saveClient(clientData) {
  return executeWithLock(async () => {
    await ensureClientsMasterInitialized();
    await backupBeforeWrite();

    const workbook = await loadWorkbook();
    const sheet = workbook.getWorksheet('Clients');

    const clientId = clientData.clientId || await generateNextClientId();
    let targetRow = null;

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      if (desanitize(row.getCell(1).value) === clientId) {
        targetRow = row;
      }
    });

    const payload = {
      clientId,
      clientName: sanitize(clientData.clientName || ''),
      contactPerson: sanitize(clientData.contactPerson || ''),
      mobile: sanitize(clientData.mobile || ''),
      email: sanitize(clientData.email || ''),
      address: sanitize(clientData.address || ''),
      state: sanitize(clientData.state || 'Uttar Pradesh'),
      stateCode: sanitize(clientData.stateCode || '09'),
      gstin: sanitize(clientData.gstin || ''),
      createdDate: clientData.createdDate || new Date().toISOString().slice(0, 10)
    };

    if (targetRow) {
      SCHEMAS.Clients.forEach((col, idx) => {
        targetRow.getCell(idx + 1).value = payload[col.key];
      });
      targetRow.commit();
    } else {
      sheet.addRow(objectToRowArray('Clients', payload));
    }

    await workbook.xlsx.writeFile(CLIENTS_MASTER_PATH);
    return payload;
  });
}

async function createInvoice(invoiceData) {
  return executeWithLock(async () => {
    await ensureClientsMasterInitialized();
    await backupBeforeWrite();

    const workbook = await loadWorkbook();
    const invoicesSheet = workbook.getWorksheet('Invoices');
    const itemsSheet = workbook.getWorksheet('InvoiceItems');

    const invoiceNo = invoiceData.invoiceNo || await generateNextInvoiceNo();
    const publicToken = crypto.randomBytes(16).toString('hex');

    const rawItems = Array.isArray(invoiceData.items) ? invoiceData.items : [];
    let taxableTotalPaise = 0;
    let taxTotalPaise = 0;
    let grandTotalPaise = 0;

    const processedItems = rawItems.map((item, index) => {
      const qty = Number(item.qty) || 1;
      const ratePaise = toPaise(item.rate);
      const taxableValuePaise = Math.round(qty * ratePaise);
      const taxPercent = Number(item.taxPercent) || 0;
      const taxAmountPaise = Math.round((taxableValuePaise * taxPercent) / 100);
      const lineTotalPaise = taxableValuePaise + taxAmountPaise;

      taxableTotalPaise += taxableValuePaise;
      taxTotalPaise += taxAmountPaise;
      grandTotalPaise += lineTotalPaise;

      return {
        invoiceNo,
        srNo: index + 1,
        description: sanitize(item.description || ''),
        hsnSac: sanitize(item.hsnSac || ''),
        qty,
        unit: sanitize(item.unit || 'NOS'),
        rate: roundToTwo(fromPaise(ratePaise)),
        taxableValue: roundToTwo(fromPaise(taxableValuePaise)),
        taxPercent,
        taxAmount: roundToTwo(fromPaise(taxAmountPaise)),
        lineTotal: roundToTwo(fromPaise(lineTotalPaise))
      };
    });

    const invoiceRow = {
      invoiceNo,
      clientId: sanitize(invoiceData.clientId || ''),
      invoiceDate: invoiceData.invoiceDate || new Date().toISOString().slice(0, 10),
      dueDate: invoiceData.dueDate || new Date(Date.now() + 7 * 86400000).toISOString().slice(0, 10),
      taxableTotal: roundToTwo(fromPaise(taxableTotalPaise)),
      taxTotal: roundToTwo(fromPaise(taxTotalPaise)),
      grandTotal: roundToTwo(fromPaise(grandTotalPaise)),
      status: 'Unpaid',
      publicToken,
      notes: sanitize(invoiceData.notes || ''),
      challanNo: sanitize(invoiceData.challanNo || ''),
      challanDate: sanitize(invoiceData.challanDate || ''),
      ewayBillNo: sanitize(invoiceData.ewayBillNo || ''),
      transport: sanitize(invoiceData.transport || ''),
      transportId: sanitize(invoiceData.transportId || ''),
      planName: sanitize(invoiceData.planName || 'Annual Cloud & Maintenance Plan'),
      serviceStartDate: sanitize(invoiceData.serviceStartDate || invoiceData.invoiceDate || new Date().toISOString().slice(0, 10)),
      serviceExpiryDate: sanitize(invoiceData.serviceExpiryDate || new Date(Date.now() + 365 * 86400000).toISOString().slice(0, 10)),
      renewalDueDate: sanitize(invoiceData.renewalDueDate || invoiceData.dueDate || new Date(Date.now() + 7 * 86400000).toISOString().slice(0, 10))
    };

    invoicesSheet.addRow(objectToRowArray('Invoices', invoiceRow));

    processedItems.forEach(item => {
      itemsSheet.addRow(objectToRowArray('InvoiceItems', item));
    });

    await workbook.xlsx.writeFile(CLIENTS_MASTER_PATH);

    return {
      ...invoiceRow,
      items: processedItems,
      totalPaid: 0,
      balanceDue: invoiceRow.grandTotal
    };
  });
}

async function updateInvoice(invoiceNo, invoiceData) {
  return executeWithLock(async () => {
    await ensureClientsMasterInitialized();
    await backupBeforeWrite();

    const workbook = await loadWorkbook();
    const invoicesSheet = workbook.getWorksheet('Invoices');
    const itemsSheet = workbook.getWorksheet('InvoiceItems');
    const paymentsSheet = workbook.getWorksheet('Payments');

    // Find existing invoice row
    let targetRowNumber = -1;
    let existingInvoice = null;
    invoicesSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const rowNo = row.getCell(1).value;
      if (String(rowNo).trim() === String(invoiceNo).trim()) {
        targetRowNumber = rowNumber;
        existingInvoice = rowToObject('Invoices', row.values);
      }
    });

    if (!existingInvoice || targetRowNumber === -1) {
      throw new Error(`Invoice ${invoiceNo} not found.`);
    }

    // Process new items and recalculate totals
    const rawItems = Array.isArray(invoiceData.items) ? invoiceData.items : [];
    let taxableTotalPaise = 0;
    let taxTotalPaise = 0;
    let grandTotalPaise = 0;

    const processedItems = rawItems.map((item, index) => {
      const qty = Number(item.qty) || 1;
      const ratePaise = toPaise(item.rate);
      const taxableValuePaise = Math.round(qty * ratePaise);
      const taxPercent = Number(item.taxPercent) || 0;
      const taxAmountPaise = Math.round((taxableValuePaise * taxPercent) / 100);
      const lineTotalPaise = taxableValuePaise + taxAmountPaise;

      taxableTotalPaise += taxableValuePaise;
      taxTotalPaise += taxAmountPaise;
      grandTotalPaise += lineTotalPaise;

      return {
        invoiceNo,
        srNo: index + 1,
        description: sanitize(item.description || ''),
        hsnSac: sanitize(item.hsnSac || ''),
        qty,
        unit: sanitize(item.unit || 'NOS'),
        rate: roundToTwo(fromPaise(ratePaise)),
        taxableValue: roundToTwo(fromPaise(taxableValuePaise)),
        taxPercent,
        taxAmount: roundToTwo(fromPaise(taxAmountPaise)),
        lineTotal: roundToTwo(fromPaise(lineTotalPaise))
      };
    });

    // Compute existing payments to accurately update status and balance
    let totalPaidPaise = 0;
    paymentsSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const rowInvNo = row.getCell(2).value;
      if (String(rowInvNo).trim() === String(invoiceNo).trim()) {
        const amt = row.getCell(3).value;
        totalPaidPaise += toPaise(amt);
      }
    });

    const grandTotal = roundToTwo(fromPaise(grandTotalPaise));
    const totalPaid = roundToTwo(fromPaise(totalPaidPaise));
    const balanceDue = Math.max(0, roundToTwo(fromPaise(grandTotalPaise - totalPaidPaise)));

    let status = 'Unpaid';
    if (balanceDue === 0 && grandTotal > 0) {
      status = 'Paid';
    } else if (totalPaid > 0 && balanceDue > 0) {
      status = 'Partially Paid';
    }

    const updatedInvoiceRow = {
      invoiceNo,
      clientId: sanitize(invoiceData.clientId || existingInvoice.clientId || ''),
      invoiceDate: invoiceData.invoiceDate || existingInvoice.invoiceDate,
      dueDate: invoiceData.dueDate || existingInvoice.dueDate,
      taxableTotal: roundToTwo(fromPaise(taxableTotalPaise)),
      taxTotal: roundToTwo(fromPaise(taxTotalPaise)),
      grandTotal,
      status,
      publicToken: existingInvoice.publicToken || crypto.randomBytes(16).toString('hex'),
      notes: sanitize(invoiceData.notes !== undefined ? invoiceData.notes : (existingInvoice.notes || '')),
      challanNo: sanitize(invoiceData.challanNo !== undefined ? invoiceData.challanNo : (existingInvoice.challanNo || '')),
      challanDate: sanitize(invoiceData.challanDate !== undefined ? invoiceData.challanDate : (existingInvoice.challanDate || '')),
      ewayBillNo: sanitize(invoiceData.ewayBillNo !== undefined ? invoiceData.ewayBillNo : (existingInvoice.ewayBillNo || '')),
      transport: sanitize(invoiceData.transport !== undefined ? invoiceData.transport : (existingInvoice.transport || '')),
      transportId: sanitize(invoiceData.transportId !== undefined ? invoiceData.transportId : (existingInvoice.transportId || '')),
      planName: sanitize(invoiceData.planName || existingInvoice.planName || 'Annual Cloud & Maintenance Plan'),
      serviceStartDate: sanitize(invoiceData.serviceStartDate || existingInvoice.serviceStartDate || ''),
      serviceExpiryDate: sanitize(invoiceData.serviceExpiryDate || existingInvoice.serviceExpiryDate || ''),
      renewalDueDate: sanitize(invoiceData.renewalDueDate || existingInvoice.renewalDueDate || '')
    };

    // Update row in Invoices sheet
    const rowValues = objectToRowArray('Invoices', updatedInvoiceRow);
    const targetRow = invoicesSheet.getRow(targetRowNumber);
    for (let c = 0; c < rowValues.length; c++) {
      targetRow.getCell(c + 1).value = rowValues[c];
    }
    targetRow.commit();

    // Replace items in InvoiceItems sheet for this invoice
    const rowsToDelete = [];
    itemsSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const itemInvNo = row.getCell(1).value;
      if (String(itemInvNo).trim() === String(invoiceNo).trim()) {
        rowsToDelete.push(rowNumber);
      }
    });
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      itemsSheet.spliceRows(rowsToDelete[i], 1);
    }
    processedItems.forEach(item => {
      itemsSheet.addRow(objectToRowArray('InvoiceItems', item));
    });

    await workbook.xlsx.writeFile(CLIENTS_MASTER_PATH);

    return {
      ...updatedInvoiceRow,
      items: processedItems,
      totalPaid,
      balanceDue
    };
  });
}

async function addPayment(invoiceNo, paymentData) {
  return executeWithLock(async () => {
    await ensureClientsMasterInitialized();
    await backupBeforeWrite();

    const workbook = await loadWorkbook();
    const invoicesSheet = workbook.getWorksheet('Invoices');
    const paymentsSheet = workbook.getWorksheet('Payments');

    let invoiceRow = null;
    let currentGrandTotal = 0;

    invoicesSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      if (desanitize(row.getCell(1).value) === invoiceNo) {
        invoiceRow = row;
        currentGrandTotal = Number(row.getCell(7).value) || 0;
      }
    });

    if (!invoiceRow) {
      throw new Error(`Invoice ${invoiceNo} not found.`);
    }

    let existingPaidPaise = 0;
    paymentsSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      if (desanitize(row.getCell(2).value) === invoiceNo) {
        existingPaidPaise += toPaise(row.getCell(3).value);
      }
    });

    const grandTotalPaise = toPaise(currentGrandTotal);
    const newAmountPaise = toPaise(paymentData.amountPaid);

    if (newAmountPaise <= 0) {
      throw new Error('Payment amount must be greater than zero.');
    }

    if (existingPaidPaise + newAmountPaise > grandTotalPaise) {
      const remainingPaise = Math.max(0, grandTotalPaise - existingPaidPaise);
      throw new Error(
        `Total paid amount cannot exceed Grand Total. Maximum allowed payment is ₹${fromPaise(remainingPaise)}.`
      );
    }

    const paymentId = `PAY-${new Date().toISOString().replace(/[-:T.]/g, '').slice(0, 8)}-${crypto.randomBytes(3).toString('hex').toUpperCase()}`;

    const newPayment = {
      paymentId,
      invoiceNo,
      amountPaid: roundToTwo(fromPaise(newAmountPaise)),
      paymentDate: paymentData.paymentDate || new Date().toISOString().slice(0, 10),
      mode: sanitize(paymentData.mode || 'UPI'),
      note: sanitize(paymentData.note || '')
    };

    paymentsSheet.addRow(objectToRowArray('Payments', newPayment));

    const totalPaidPaise = existingPaidPaise + newAmountPaise;
    let newStatus = 'Partially Paid';
    if (totalPaidPaise >= grandTotalPaise) {
      newStatus = 'Paid';
    }

    invoiceRow.getCell(8).value = newStatus;
    invoiceRow.commit();

    await workbook.xlsx.writeFile(CLIENTS_MASTER_PATH);

    const balanceDue = roundToTwo(fromPaise(Math.max(0, grandTotalPaise - totalPaidPaise)));

    return {
      payment: newPayment,
      invoiceNo,
      newStatus,
      totalPaid: roundToTwo(fromPaise(totalPaidPaise)),
      balanceDue
    };
  });
}

module.exports = {
  ensureClientsMasterInitialized,
  getAllClients,
  getClientById,
  saveClient,
  getAllInvoices,
  getInvoiceByNo,
  getInvoiceByToken,
  createInvoice,
  updateInvoice,
  addPayment,
  getPaymentsForInvoice,
  generateNextInvoiceNo,
  generateNextClientId,
  toPaise,
  fromPaise,
  roundToTwo
};
