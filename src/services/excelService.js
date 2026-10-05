const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

// Determine storage path (Support local server, VPS, and /tmp for serverless)
const defaultDataDir = process.env.VERCEL
  ? path.join('/tmp', 'yogteck-data')
  : path.join(__dirname, '../../data');

const dataDir = process.env.DATA_DIR || defaultDataDir;
const excelFilePath = process.env.EXCEL_FILE_PATH || path.join(dataDir, 'enquiries.xlsx');
const jsonBackupPath = path.join(dataDir, 'enquiries_backup.jsonl');

// Ensure directory exists
function ensureDataDirectory() {
  const dir = path.dirname(excelFilePath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
}

// Mutex promise queue to ensure atomic sequential file writes
let writeQueue = Promise.resolve();

function getIstDateTime(dateObj = new Date()) {
  const optionsDate = {
    timeZone: 'Asia/Kolkata',
    day: '2-digit',
    month: '2-digit',
    year: 'numeric'
  };
  const optionsTime = {
    timeZone: 'Asia/Kolkata',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: true
  };

  const formattedDate = dateObj.toLocaleDateString('en-GB', optionsDate); // DD/MM/YYYY
  const formattedTime = dateObj.toLocaleTimeString('en-IN', optionsTime); // hh:mm:ss AM/PM

  return {
    dateIST: formattedDate.replace(/\//g, '-'),
    timeIST: formattedTime,
    isoTimestamp: dateObj.toISOString()
  };
}

const COLUMNS = [
  { header: 'S.No.', key: 'sno', width: 8 },
  { header: 'Date (IST)', key: 'date', width: 14 },
  { header: 'Time (IST)', key: 'time', width: 16 },
  { header: 'Customer Name', key: 'name', width: 24 },
  { header: 'Company Name', key: 'companyName', width: 26 },
  { header: 'Mobile / WhatsApp', key: 'phone', width: 18 },
  { header: 'Email Address', key: 'email', width: 28 },
  { header: 'Required Service', key: 'serviceType', width: 30 },
  { header: 'Project Requirements / Message', key: 'message', width: 45 },
  { header: 'IP Address', key: 'ip', width: 18 },
  { header: 'Timestamp (ISO)', key: 'timestamp', width: 26 }
];

function applyHeaderStyles(worksheet) {
  const headerRow = worksheet.getRow(1);
  headerRow.height = 30;
  
  headerRow.eachCell((cell) => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF0A1330' } // Navy Brand Color
    };
    cell.font = {
      name: 'Calibri',
      size: 11,
      bold: true,
      color: { argb: 'FFFFFFFF' }
    };
    cell.alignment = {
      vertical: 'middle',
      horizontal: 'center',
      wrapText: true
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFFF9A1F' } },
      bottom: { style: 'medium', color: { argb: 'FFFF9A1F' } },
      left: { style: 'thin', color: { argb: 'FF1E293B' } },
      right: { style: 'thin', color: { argb: 'FF1E293B' } }
    };
  });
}

function applyDataRowStyles(row, rowNumber) {
  row.height = 24;
  const isEven = rowNumber % 2 === 0;

  row.eachCell((cell, colNumber) => {
    cell.font = {
      name: 'Calibri',
      size: 10,
      color: { argb: 'FF0F172A' }
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
    };

    if (isEven) {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF8FAFC' }
      };
    }

    // Alignments
    if (colNumber === 1 || colNumber === 2 || colNumber === 3) {
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    } else if (colNumber === 9) { // Message
      cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    } else {
      cell.alignment = { vertical: 'middle', horizontal: 'left' };
    }
  });
}

async function appendEnquiryToExcel(enquiryData, clientIp = '') {
  return new Promise((resolve, reject) => {
    writeQueue = writeQueue.then(async () => {
      try {
        ensureDataDirectory();
        const { dateIST, timeIST, isoTimestamp } = getIstDateTime();

        const workbook = new ExcelJS.Workbook();
        let worksheet;

        if (fs.existsSync(excelFilePath)) {
          await workbook.xlsx.readFile(excelFilePath);
          worksheet = workbook.getWorksheet('Enquiries') || workbook.worksheets[0];
          if (!worksheet) {
            worksheet = workbook.addWorksheet('Enquiries');
            worksheet.columns = COLUMNS;
            applyHeaderStyles(worksheet);
          }
        } else {
          worksheet = workbook.addWorksheet('Enquiries');
          worksheet.columns = COLUMNS;
          applyHeaderStyles(worksheet);
        }

        // Determine next Serial Number
        const rowCount = worksheet.rowCount; // includes header
        const sno = rowCount; // Since row 1 is header, next row index is rowCount + 1, sno is rowCount

        const newRowData = {
          sno: sno,
          date: dateIST,
          time: timeIST,
          name: enquiryData.name || '',
          companyName: enquiryData.companyName || 'N/A',
          phone: enquiryData.phone || '',
          email: enquiryData.email || '',
          serviceType: enquiryData.rackType || enquiryData.serviceType || 'Website & Digital Solutions',
          message: enquiryData.message || '',
          ip: clientIp || 'N/A',
          timestamp: isoTimestamp
        };

        const addedRow = worksheet.addRow(newRowData);
        applyDataRowStyles(addedRow, rowCount + 1);

        // Adjust column widths if needed
        worksheet.columns.forEach(column => {
          let maxLength = 0;
          column.eachCell({ includeEmpty: false }, (cell) => {
            const cellValue = cell.value ? cell.value.toString() : '';
            if (cellValue.length > maxLength) {
              maxLength = cellValue.length;
            }
          });
          column.width = Math.min(Math.max(maxLength + 3, column.width || 12), 60);
        });

        await workbook.xlsx.writeFile(excelFilePath);
        console.log(`[EXCEL LOGGED] Successfully appended enquiry #${sno} (${enquiryData.name}) to ${excelFilePath}`);

        // Also write to JSON Lines backup for 100% redundancy
        try {
          const jsonRow = JSON.stringify({ ...newRowData, createdAt: isoTimestamp }) + '\n';
          fs.appendFileSync(jsonBackupPath, jsonRow, 'utf8');
        } catch (jsonErr) {
          console.warn('JSON backup write warning:', jsonErr);
        }

        resolve({ success: true, filePath: excelFilePath, sno, dateIST, timeIST });
      } catch (err) {
        console.error('Failed to append enquiry to Excel:', err);
        reject(err);
      }
    });
  });
}

function parseCellValue(val) {
  if (val === null || val === undefined) return '';
  if (typeof val === 'object') {
    if (val.text) return val.text;
    if (val.result !== undefined) return val.result;
    return JSON.stringify(val);
  }
  return String(val).trim();
}

async function readAllEnquiriesFromExcel() {
  try {
    ensureDataDirectory();

    if (fs.existsSync(excelFilePath)) {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(excelFilePath);
      const worksheet = workbook.getWorksheet('Enquiries') || workbook.worksheets[0];

      if (worksheet && worksheet.rowCount > 1) {
        const enquiries = [];

        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1) return; // Skip header row

          const snoVal = parseCellValue(row.getCell(1).value);
          const dateVal = parseCellValue(row.getCell(2).value);
          const timeVal = parseCellValue(row.getCell(3).value);
          const nameVal = parseCellValue(row.getCell(4).value);
          const companyVal = parseCellValue(row.getCell(5).value);
          const phoneVal = parseCellValue(row.getCell(6).value);
          const emailVal = parseCellValue(row.getCell(7).value);
          const serviceVal = parseCellValue(row.getCell(8).value);
          const messageVal = parseCellValue(row.getCell(9).value);
          const ipVal = parseCellValue(row.getCell(10).value);
          const timestampVal = parseCellValue(row.getCell(11).value);

          // Skip completely empty rows
          if (!nameVal && !phoneVal && !emailVal && !messageVal) return;

          enquiries.push({
            sno: Number(snoVal) || rowNumber - 1,
            date: dateVal,
            time: timeVal,
            name: nameVal,
            companyName: companyVal === 'N/A' ? '' : companyVal,
            phone: phoneVal,
            email: emailVal,
            serviceType: serviceVal || 'Website & Digital Solutions',
            message: messageVal,
            ip: ipVal,
            timestamp: timestampVal || ''
          });
        });

        // Sort descending: newest inquiry first (by timestamp descending or sno descending)
        enquiries.sort((a, b) => {
          if (a.timestamp && b.timestamp) {
            const timeA = new Date(a.timestamp).getTime();
            const timeB = new Date(b.timestamp).getTime();
            if (!isNaN(timeA) && !isNaN(timeB)) {
              return timeB - timeA;
            }
          }
          return b.sno - a.sno;
        });

        return {
          total: enquiries.length,
          inquiries: enquiries,
          source: 'excel'
        };
      }
    }

    // Fallback: Check JSONL backup
    if (fs.existsSync(jsonBackupPath)) {
      const content = fs.readFileSync(jsonBackupPath, 'utf8');
      const lines = content.split('\n').map(l => l.trim()).filter(Boolean);
      const enquiries = [];

      lines.forEach((line, idx) => {
        try {
          const parsed = JSON.parse(line);
          enquiries.push({
            sno: parsed.sno || idx + 1,
            date: parsed.date || '',
            time: parsed.time || '',
            name: parsed.name || '',
            companyName: parsed.companyName === 'N/A' ? '' : (parsed.companyName || ''),
            phone: parsed.phone || '',
            email: parsed.email || '',
            serviceType: parsed.serviceType || 'Website & Digital Solutions',
            message: parsed.message || '',
            ip: parsed.ip || '',
            timestamp: parsed.timestamp || parsed.createdAt || ''
          });
        } catch (e) {}
      });

      // Sort descending
      enquiries.sort((a, b) => {
        if (a.timestamp && b.timestamp) {
          const timeA = new Date(a.timestamp).getTime();
          const timeB = new Date(b.timestamp).getTime();
          if (!isNaN(timeA) && !isNaN(timeB)) {
            return timeB - timeA;
          }
        }
        return b.sno - a.sno;
      });

      return {
        total: enquiries.length,
        inquiries: enquiries,
        source: 'json_backup'
      };
    }

    return {
      total: 0,
      inquiries: [],
      source: 'none'
    };
  } catch (error) {
    console.error('Error reading enquiries from Excel:', error);
    return {
      total: 0,
      inquiries: [],
      error: error.message
    };
  }
}

function getExcelFilePath() {
  return excelFilePath;
}

function existsExcelFile() {
  return fs.existsSync(excelFilePath);
}

module.exports = {
  appendEnquiryToExcel,
  readAllEnquiriesFromExcel,
  getExcelFilePath,
  existsExcelFile,
  getIstDateTime
};
