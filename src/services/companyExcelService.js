const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

const defaultDataDir = process.env.VERCEL
  ? path.join('/tmp', 'yogteck-data')
  : path.join(__dirname, '..', '..', 'data');

const seedDataDir = path.join(__dirname, '..', '..', 'data');
const dataDir = process.env.DATA_DIR || defaultDataDir;
const COMPANY_FILE_PATH = path.join(dataDir, 'company_master.xlsx');
const SEED_COMPANY_FILE = path.join(seedDataDir, 'company_master.xlsx');

const DEFAULT_IMPORTANT_NOTES = `1) भुगतान और सर्विस रिन्यूअल / Payment & Service Renewal
यदि भुगतान या रिन्यूअल में देरी होती है, तो सर्विसेज़ की अवधि समाप्त (Expire) हो जाने के कारण आपका प्रोजेक्ट/वेबसाइट कभी भी बंद हो सकता है। इसलिए कृपया सर्विसेज़ को निर्बाध रूप से चालू रखने के लिए समय पर भुगतान या रिन्यूअल कर लें। देरी की स्थिति में किसी भी प्रकार के नुकसान या डेटा/सेवा बाधा के लिए कंपनी उत्तरदायी नहीं होगी।
Delay in payment or renewal may cause your services to expire and your project/website to go offline at any time. Please pay or renew on time to keep services active. The company will not be responsible for any interruption or loss caused by delayed payment.

2) सब्सक्रिप्शन की जानकारी / Subscription & Plan Details
अपने सब्सक्रिप्शन से जुड़ी पूरी जानकारी देखने के लिए इस इनवॉइस पर दिए गए QR कोड को स्कैन करें। वहाँ आपको पूरा इनवॉइस दिखेगा, जिसमें आपके सभी प्लान, उनकी शुरुआत और समाप्ति की तिथियाँ (Expiry), भुगतान का इतिहास और बाकी रकम (Balance Due) सहित सभी संबंधित विवरण उपलब्ध होंगे।
Scan the QR code on this invoice to view your full invoice, active plan details, start and expiry dates, payment history, and balance due.

3) क्लाइंट द्वारा उपलब्ध कराई जाने वाली सामग्री / Materials & External Resources to be Provided by Client
प्रोजेक्ट से जुड़ी सभी बाहरी (External) फाइलें, सेवाएँ और संसाधन क्लाइंट/पार्टी को ही उपलब्ध कराने होंगे। उदाहरण के लिए:
• पेमेंट गेटवे (Payment Gateway) का अकाउंट/क्रेडेंशइल्स क्लाइंट द्वारा उपलब्ध कराया जाएगा।
• प्रोजेक्ट से जुड़ी क्लाइंट की अपनी फाइलें, जैसे CAD ड्रॉइंग्स, लोगो, कंटेंट, फोटो और अन्य दस्तावेज़ (जैसे MMR Constructions के प्रोजेक्ट में उनकी CAD फाइलें) क्लाइंट को ही देनी होंगी।
इन सामग्रियों के अभाव या देरी से होने वाली किसी भी बाधा या विलंब की ज़िम्मेदारी कंपनी की नहीं होगी।
All external files, accounts, and resources required for the project (e.g. payment gateway credentials, CAD files, logo, content, photos) must be provided by the client; the company cannot be held responsible for delays arising from their absence.`;

const DEFAULT_COMPANY_DATA = {
  companyName: 'YOGTECK Business Solutions',
  ownerName: 'Yogeshwar',
  contact: '8299209905',
  email: 'support@yogtech.com',
  address: '302/2 मंगला विहार, सेकंड 60 फिट रोड, New PAC Line, Kanpur Nagar',
  pincode: '208015',
  state: 'Uttar Pradesh',
  stateCode: '09',
  pan: 'AAHCY1234F',
  gstin: '09AAHCY1234F1Z5',
  bankName: 'ICICI Bank',
  branch: 'Kanpur Nagar',
  accountNo: '123405009876',
  ifsc: 'ICIC0001234',
  upiId: '8299209905@upi',
  logoPath: 'assets/logo/yogteck-logo.webp',
  tagline: 'Websites, Custom Software, Mobile Apps & Complete Digital Solutions',
  terms: '1. Payment due within 7 days of invoice date.\n2. Goods or Services once rendered will not be refunded.\n3. All disputes subject to Kanpur Jurisdiction.',
  gstEnabled: 'Yes',
  importantNotes: DEFAULT_IMPORTANT_NOTES
};

const COMPANY_COLUMNS = [
  { header: 'Company Name', key: 'companyName', width: 30 },
  { header: 'Owner Name', key: 'ownerName', width: 22 },
  { header: 'Contact', key: 'contact', width: 16 },
  { header: 'Email', key: 'email', width: 28 },
  { header: 'Address', key: 'address', width: 50 },
  { header: 'Pincode', key: 'pincode', width: 12 },
  { header: 'State', key: 'state', width: 18 },
  { header: 'State Code', key: 'stateCode', width: 12 },
  { header: 'PAN', key: 'pan', width: 16 },
  { header: 'GSTIN', key: 'gstin', width: 20 },
  { header: 'Bank Name', key: 'bankName', width: 20 },
  { header: 'Branch', key: 'branch', width: 20 },
  { header: 'Account No', key: 'accountNo', width: 22 },
  { header: 'IFSC', key: 'ifsc', width: 16 },
  { header: 'UPI ID', key: 'upiId', width: 24 },
  { header: 'Logo Path', key: 'logoPath', width: 30 },
  { header: 'Tagline', key: 'tagline', width: 45 },
  { header: 'Terms', key: 'terms', width: 50 },
  { header: 'GST Enabled', key: 'gstEnabled', width: 14 },
  { header: 'Important Notes', key: 'importantNotes', width: 60 }
];

function sanitizeForExcel(val) {
  if (typeof val !== 'string') return val;
  const trimmed = val.trim();
  if (/^[=+\-@]/.test(trimmed)) {
    return `'${trimmed}`;
  }
  return trimmed;
}

function desanitizeFromExcel(val) {
  if (val == null) return '';
  if (typeof val === 'object' && val.text) val = val.text;
  const str = String(val).trim();
  if (str.startsWith("'") && /^[=+\-@]/.test(str.slice(1))) {
    return str.slice(1);
  }
  return str;
}

async function ensureCompanyMasterInitialized() {
  if (!fs.existsSync(dataDir)) {
    try {
      fs.mkdirSync(dataDir, { recursive: true });
    } catch (e) {
      console.warn('[COMPANY DIR MKDIR]', e.message);
    }
  }

  // If in Vercel/tmp and file not in /tmp yet, copy from seed if exists
  if (!fs.existsSync(COMPANY_FILE_PATH) && fs.existsSync(SEED_COMPANY_FILE)) {
    try {
      fs.copyFileSync(SEED_COMPANY_FILE, COMPANY_FILE_PATH);
      return;
    } catch (e) {
      console.warn('Could not copy seed company file:', e.message);
    }
  }

  if (fs.existsSync(COMPANY_FILE_PATH)) {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(COMPANY_FILE_PATH);
      const sheet = workbook.getWorksheet('Company') || workbook.worksheets[0];
      if (sheet) {
        let hasNotesCol = false;
        sheet.getRow(1).eachCell((cell) => {
          if (String(cell.value || '').toLowerCase().includes('important notes')) {
            hasNotesCol = true;
          }
        });

        if (!hasNotesCol) {
          const notesColIdx = COMPANY_COLUMNS.length; // Col 20
          sheet.getRow(1).getCell(notesColIdx).value = 'Important Notes';
          sheet.getRow(1).getCell(notesColIdx).font = { bold: true, color: { argb: 'FFFFFFFF' } };
          sheet.getRow(1).getCell(notesColIdx).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF0A1330' }
          };

          if (sheet.rowCount >= 2) {
            sheet.getRow(2).getCell(notesColIdx).value = sanitizeForExcel(DEFAULT_IMPORTANT_NOTES);
          }
          await workbook.xlsx.writeFile(COMPANY_FILE_PATH);
        }
      }
    } catch (e) {
      console.warn('[COMPANY MASTER UPGRADE]', e.message);
    }
    return;
  }

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Company');
  sheet.columns = COMPANY_COLUMNS;

  // Header style
  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF0A1330' }
  };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

  // Add default row
  const rowValues = COMPANY_COLUMNS.map(col => sanitizeForExcel(DEFAULT_COMPANY_DATA[col.key] || ''));
  sheet.addRow(rowValues);

  await workbook.xlsx.writeFile(COMPANY_FILE_PATH);
}

async function getCompanyMaster() {
  await ensureCompanyMasterInitialized();

  if (!fs.existsSync(COMPANY_FILE_PATH)) {
    return { ...DEFAULT_COMPANY_DATA };
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(COMPANY_FILE_PATH);
  const sheet = workbook.getWorksheet('Company') || workbook.worksheets[0];

  if (!sheet || sheet.rowCount < 2) {
    return { ...DEFAULT_COMPANY_DATA };
  }

  const row = sheet.getRow(2);
  const data = {};
  COMPANY_COLUMNS.forEach((col, idx) => {
    const rawVal = row.getCell(idx + 1).value;
    data[col.key] = desanitizeFromExcel(rawVal);
  });

  if (!data.importantNotes) {
    data.importantNotes = DEFAULT_IMPORTANT_NOTES;
  }

  return {
    ...DEFAULT_COMPANY_DATA,
    ...data
  };
}

async function updateCompanyMaster(updatedData) {
  await ensureCompanyMasterInitialized();

  const workbook = new ExcelJS.Workbook();
  if (fs.existsSync(COMPANY_FILE_PATH)) {
    await workbook.xlsx.readFile(COMPANY_FILE_PATH);
  }
  let sheet = workbook.getWorksheet('Company');
  if (!sheet) {
    sheet = workbook.addWorksheet('Company');
    sheet.columns = COMPANY_COLUMNS;
  }

  const merged = {
    ...DEFAULT_COMPANY_DATA,
    ...updatedData
  };

  const row = sheet.getRow(2);
  COMPANY_COLUMNS.forEach((col, idx) => {
    const val = merged[col.key];
    row.getCell(idx + 1).value = sanitizeForExcel(val != null ? String(val) : '');
  });
  row.commit();

  await workbook.xlsx.writeFile(COMPANY_FILE_PATH);
  return merged;
}

module.exports = {
  getCompanyMaster,
  updateCompanyMaster,
  ensureCompanyMasterInitialized,
  DEFAULT_IMPORTANT_NOTES
};
