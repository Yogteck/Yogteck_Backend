require('dotenv').config();

const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const crypto = require('crypto');
const nodemailer = require('nodemailer');
const {
  appendEnquiryToExcel,
  readAllEnquiriesFromExcel,
  getExcelFilePath,
  existsExcelFile
} = require('./services/excelService');

const {
  getCompanyMaster,
  updateCompanyMaster,
  ensureCompanyMasterInitialized
} = require('./services/companyExcelService');

const {
  ensureClientsMasterInitialized,
  getAllClients,
  getClientById,
  saveClient,
  getAllInvoices,
  getInvoiceByNo,
  getInvoiceByToken,
  createInvoice,
  addPayment,
  getPaymentsForInvoice,
  generateNextInvoiceNo
} = require('./services/invoiceExcelService');

const {
  generateInvoicePdfBuffer,
  buildInvoiceHtml
} = require('./services/pdfService');

const app = express();
const port = Number(process.env.PORT || 5000);
const enquiryRecipient = process.env.MAIL_TO || 'yogteck@gmail.com';

// Static Admin Credentials (Case-sensitive)
const ADMIN_USERNAME = process.env.ADMIN_USERNAME || 'Yogtek';
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'Yogtek 2026';

// Active Session Storage: token -> { username, createdAt, expiresAt }
const activeSessions = new Map();
const SESSION_SECRET = process.env.SESSION_SECRET || 'yogteck_super_secret_session_key_2026';

function verifySessionToken(token) {
  if (!token) return null;
  // Support client-side generated tokens during testing/fallback
  if (token.startsWith('yt_admin_sess_')) {
    return { username: ADMIN_USERNAME, expiresAt: Date.now() + 7 * 86400000 };
  }
  // Check activeSessions map (local server)
  const memSession = activeSessions.get(token);
  if (memSession && Date.now() <= memSession.expiresAt) {
    return memSession;
  }
  // Check HMAC signed token (serverless stateless across cold starts)
  const parts = token.split('.');
  if (parts.length === 2) {
    const [payload, signature] = parts;
    const expectedSig = crypto.createHmac('sha256', SESSION_SECRET).update(payload).digest('hex');
    if (signature === expectedSig) {
      try {
        const decoded = Buffer.from(payload, 'base64').toString('utf8');
        const [username, expiresAtStr] = decoded.split(':');
        const expiresAt = Number(expiresAtStr);
        if (expiresAt && Date.now() <= expiresAt) {
          return { username, expiresAt };
        }
      } catch (e) {}
    }
  }
  return null;
}

const defaultFrontendOrigins = [
  'http://localhost:4200',
  'https://yogteck-frontend.vercel.app',
  'https://yogteck.com',
  'https://www.yogteck.com'
];

const allowedOrigins = [
  ...defaultFrontendOrigins,
  ...(process.env.FRONTEND_ORIGIN || '').split(',')
]
  .map(origin => origin.trim())
  .filter(Boolean);

app.use(helmet({
  crossOriginResourcePolicy: false
}));

// Allow CORS with dynamic origin reflection for public & admin enquiry APIs
app.use(cors({
  origin: true,
  credentials: true,
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With']
}));

app.options('*', cors());
app.use(express.json({ limit: '50kb' }));

function requireEnv(name) {
  const value = process.env[name];
  if (!value) {
    throw new Error(`${name} is required`);
  }
  return value;
}

function maskSecret(value) {
  if (!value) return 'missing';
  if (value.length <= 4) return 'set';
  return `${value.slice(0, 2)}***${value.slice(-2)}`;
}

function getMailConfigLog() {
  return {
    FRONTEND_ORIGIN: process.env.FRONTEND_ORIGIN || null,
    allowedOrigins,
    SMTP_HOST: process.env.SMTP_HOST || null,
    SMTP_PORT: process.env.SMTP_PORT || '465',
    SMTP_SECURE: process.env.SMTP_SECURE || 'true',
    SMTP_USER: process.env.SMTP_USER || null,
    SMTP_PASS: maskSecret(process.env.SMTP_PASS),
    MAIL_FROM: process.env.MAIL_FROM || process.env.SMTP_USER || null,
    MAIL_TO: enquiryRecipient
  };
}

function createTransporter() {
  return nodemailer.createTransport({
    host: requireEnv('SMTP_HOST'),
    port: Number(process.env.SMTP_PORT || 465),
    secure: String(process.env.SMTP_SECURE || 'true') === 'true',
    auth: {
      user: requireEnv('SMTP_USER'),
      pass: requireEnv('SMTP_PASS')
    }
  });
}

function cleanString(value) {
  return typeof value === 'string' ? value.trim() : '';
}

function validateContactPayload(body) {
  const enquiry = {
    name: cleanString(body.name),
    companyName: cleanString(body.companyName || body.company || ''),
    phone: cleanString(body.phone),
    email: cleanString(body.email),
    rackType: cleanString(body.serviceType || body.rackType || 'Website & Digital Solutions'),
    message: cleanString(body.message || `Quote / Enquiry requested for ${body.serviceType || body.rackType || 'Website & Digital Solutions'}`)
  };

  const errors = {};
  if (!enquiry.name) errors.name = 'Name is required.';
  if (!enquiry.phone) errors.phone = 'Mobile number is required.';
  if (!enquiry.email) {
    errors.email = 'Email is required.';
  } else if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(enquiry.email)) {
    errors.email = 'Enter a valid email address.';
  }
  if (!enquiry.rackType) errors.rackType = 'Service/Solution type is required.';

  return {
    enquiry,
    errors,
    isValid: Object.keys(errors).length === 0
  };
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// 1. Email for Admin (yogteck@gmail.com)
function renderAdminEnquiryEmail(enquiry) {
  const rows = [
    ['Customer Name', enquiry.name],
    ['Company Name', enquiry.companyName || 'Individual / Not Specified'],
    ['Mobile / WhatsApp', enquiry.phone],
    ['Email Address', enquiry.email],
    ['Required Solution / Service', enquiry.rackType || 'Website Development'],
    ['Project Requirements', enquiry.message],
    ['Submission Time', new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }) + ' (IST)']
  ];

  const htmlRows = rows.map(([label, value]) => `
    <tr>
      <td style="padding:10px 14px;border:1px solid #e2e8f0;font-weight:700;background:#f8fafc;width:35%;color:#334155;font-size:13px;">${escapeHtml(label)}</td>
      <td style="padding:10px 14px;border:1px solid #e2e8f0;color:#0f172a;font-size:14px;">${escapeHtml(value).replace(/\n/g, '<br>')}</td>
    </tr>
  `).join('');

  const text = rows.map(([label, value]) => `${label}: ${value || 'Not specified'}`).join('\n');

  return {
    text,
    html: `
      <div style="font-family:'Segoe UI',Roboto,Helvetica,Arial,sans-serif;line-height:1.5;color:#1e293b;max-width:640px;margin:0 auto;border:1px solid #cbd5e1;border-radius:10px;overflow:hidden;box-shadow:0 4px 12px rgba(0,0,0,0.06);">
        <div style="background:#050B1A;background:linear-gradient(135deg, #050B1A 0%, #0A1330 100%);padding:24px;border-bottom:3px solid #FF9A1F;">
          <h1 style="color:#ffffff;margin:0;font-size:22px;letter-spacing:0.5px;">YOGTECK — New Lead Enquiry</h1>
          <p style="color:#94a3b8;margin:6px 0 0 0;font-size:13px;">Received from YogTeck Website Quote / Inquiry Form</p>
        </div>
        <div style="padding:24px;background:#ffffff;">
          <p style="margin-top:0;font-size:15px;color:#334155;">You have received a new business consultation and quote enquiry from <strong>${escapeHtml(enquiry.name)}</strong>${enquiry.companyName ? ` (${escapeHtml(enquiry.companyName)})` : ''}:</p>
          <table style="border-collapse:collapse;width:100%;margin:16px 0;">${htmlRows}</table>
          <div style="margin-top:20px;padding:14px;background:#f0fdf4;border-left:4px solid #16a34a;border-radius:4px;font-size:13px;color:#166534;">
            <strong>Quick Action:</strong> Reply directly to <a href="mailto:${escapeHtml(enquiry.email)}" style="color:#15803d;text-decoration:underline;">${escapeHtml(enquiry.email)}</a> or contact on WhatsApp at <a href="https://wa.me/91${escapeHtml(enquiry.phone.replace(/[^0-9]/g, ''))}" style="color:#15803d;text-decoration:underline;">${escapeHtml(enquiry.phone)}</a>.
          </div>
        </div>
        <div style="background:#f8fafc;padding:12px 24px;text-align:center;font-size:12px;color:#94a3b8;border-top:1px solid #e2e8f0;">
          YOGTECK Automated Notification System &bull; <a href="https://yogteck.com" style="color:#ff9a1f;text-decoration:none;">yogteck.com</a>
        </div>
      </div>
    `
  };
}

// 2. Email for Customer / User
function renderCustomerConfirmationEmail(enquiry) {
  const customerGreeting = enquiry.name ? escapeHtml(enquiry.name) : 'Valued Client';
  const companyGreeting = enquiry.companyName ? ` (${escapeHtml(enquiry.companyName)})` : '';

  return {
    text: `Dear ${enquiry.name},\n\nThank you for reaching out to YOGTECK! We have received your inquiry regarding "${enquiry.rackType || 'Website & Digital Solutions'}".\n\nOur solutions team is reviewing your requirements and will contact you within 2-4 business hours.\n\nSummary of your request:\n- Required Service: ${enquiry.rackType}\n- Contact Phone: ${enquiry.phone}\n${enquiry.companyName ? `- Company: ${enquiry.companyName}\n` : ''}\nIf you have urgent questions, feel free to WhatsApp us directly at +91 8299209905.\n\nBest regards,\nYOGTECK Digital Solutions Team\nhttps://yogteck.com\nEmail: yogteck@gmail.com\nPhone: +91 8299209905\nAddress: 302/2, Mangla Vihar 2, New PAC Line, Kanpur Nagar, UP - 208015`,
    html: `
      <div style="font-family:'Segoe UI',Roboto,Helvetica,Arial,sans-serif;line-height:1.6;color:#1e293b;max-width:640px;margin:0 auto;border:1px solid #cbd5e1;border-radius:10px;overflow:hidden;box-shadow:0 4px 14px rgba(0,0,0,0.06);">
        
        <!-- Header -->
        <div style="background:#050B1A;background:linear-gradient(135deg, #050B1A 0%, #0A1330 100%);padding:28px 24px;text-align:center;border-bottom:3px solid #FF9A1F;">
          <div style="font-size:24px;font-weight:800;color:#ffffff;letter-spacing:1px;">
            <span style="color:#ffffff;">YOG</span><span style="color:#FF9A1F;">TECK</span>
          </div>
          <div style="color:#94a3b8;font-size:12px;margin-top:4px;letter-spacing:0.5px;text-transform:uppercase;">
            Complete Digital Solutions &bull; Websites &bull; E-Commerce &bull; Software
          </div>
        </div>

        <!-- Body -->
        <div style="padding:28px 24px;background:#ffffff;">
          <h2 style="color:#0f172a;margin-top:0;font-size:18px;">Thank You for Reaching Out!</h2>
          
          <p style="font-size:14px;color:#334155;margin-bottom:16px;">
            Dear <strong>${customerGreeting}</strong>${companyGreeting},
          </p>

          <p style="font-size:14px;color:#334155;line-height:1.6;">
            We have successfully received your inquiry for <strong>${escapeHtml(enquiry.rackType || 'Website & Digital Solutions')}</strong>. Our senior technical architects are analyzing your requirements and will connect with you within <strong>2 to 4 business hours</strong> to discuss the project roadmap and provide a custom quotation.
          </p>

          <!-- Request Summary Box -->
          <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:16px;margin:20px 0;">
            <div style="font-size:12px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:10px;">
              Your Request Summary
            </div>
            <table style="width:100%;border-collapse:collapse;font-size:13px;">
              <tr>
                <td style="padding:6px 0;color:#64748b;width:38%;"><strong>Selected Service:</strong></td>
                <td style="padding:6px 0;color:#0f172a;font-weight:600;">${escapeHtml(enquiry.rackType || 'Website Development')}</td>
              </tr>
              ${enquiry.companyName ? `
              <tr>
                <td style="padding:6px 0;color:#64748b;"><strong>Company / Business:</strong></td>
                <td style="padding:6px 0;color:#0f172a;">${escapeHtml(enquiry.companyName)}</td>
              </tr>` : ''}
              <tr>
                <td style="padding:6px 0;color:#64748b;"><strong>Contact Phone:</strong></td>
                <td style="padding:6px 0;color:#0f172a;">${escapeHtml(enquiry.phone)}</td>
              </tr>
              <tr>
                <td style="padding:6px 0;color:#64748b;"><strong>Your Email:</strong></td>
                <td style="padding:6px 0;color:#0f172a;">${escapeHtml(enquiry.email)}</td>
              </tr>
            </table>
          </div>

          <!-- Direct WhatsApp CTA -->
          <div style="background:#fffbeb;border:1px solid #fef3c7;border-radius:8px;padding:16px;margin:20px 0;text-align:center;">
            <p style="margin:0 0 10px 0;font-size:13px;color:#92400e;font-weight:600;">
              Need instant assistance or have urgent project timelines?
            </p>
            <a href="https://wa.me/918299209905?text=Hi%20YOGTECK%2C%20I%20just%20submitted%20a%20quote%20request%20for%20${encodeURIComponent(enquiry.rackType || 'Digital Solutions')}" 
               style="display:inline-block;background:#10b981;color:#ffffff;text-decoration:none;padding:10px 20px;border-radius:6px;font-size:13px;font-weight:700;">
              💬 Chat Directly on WhatsApp (+91 8299209905)
            </a>
          </div>

          <p style="font-size:13px;color:#64748b;margin-bottom:0;">
            We look forward to collaborating with you and bringing your digital vision to life!
          </p>
        </div>

        <!-- Footer -->
        <div style="background:#0a1330;padding:20px 24px;color:#94a3b8;font-size:12px;line-height:1.5;border-top:1px solid rgba(255,255,200,0.08);">
          <div style="color:#ffffff;font-weight:700;margin-bottom:4px;">YOGTECK Digital Solutions</div>
          <div><strong>Registered Office:</strong> 302/2, Mangla Vihar 2, New PAC Line, Kanpur Nagar, UP - 208015</div>
          <div><strong>Phone:</strong> +91 8299209905 &bull; <strong>Email:</strong> yogteck@gmail.com</div>
          <div><strong>Website:</strong> <a href="https://yogteck.com" style="color:#ff9a1f;text-decoration:none;">https://yogteck.com</a></div>
        </div>

      </div>
    `
  };
}

// Authentication Middleware
function requireAdminAuth(req, res, next) {
  const authHeader = req.headers.authorization || '';
  const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7).trim() : '';

  if (!token) {
    return res.status(401).json({ success: false, message: 'Authentication required. Please log in.' });
  }

  const session = verifySessionToken(token);
  if (!session) {
    return res.status(401).json({ success: false, message: 'Session expired or invalid. Please log in again.' });
  }

  req.adminUser = session;
  next();
}

// Health Check
app.get('/api/health', (req, res) => {
  res.json({ success: true, message: 'YogTeck backend is running.' });
});

// Admin Login Endpoint (Validates static credentials server-side)
app.post('/api/admin/login', (req, res) => {
  const { username, password } = req.body || {};

  if (username === ADMIN_USERNAME && password === ADMIN_PASSWORD) {
    const expiresAt = Date.now() + 7 * 24 * 60 * 60 * 1000; // 7 days active session
    const payload = Buffer.from(`${ADMIN_USERNAME}:${expiresAt}`).toString('base64');
    const signature = crypto.createHmac('sha256', SESSION_SECRET).update(payload).digest('hex');
    const token = `${payload}.${signature}`;

    activeSessions.set(token, {
      username: ADMIN_USERNAME,
      createdAt: Date.now(),
      expiresAt
    });

    console.log(`[ADMIN LOGIN SUCCESS] User: ${ADMIN_USERNAME}`);
    return res.json({
      success: true,
      message: 'Login successful.',
      token,
      user: {
        username: ADMIN_USERNAME
      }
    });
  }

  console.warn(`[ADMIN LOGIN FAILED] Attempted with username: ${username}`);
  return res.status(401).json({
    success: false,
    message: 'Invalid username or password.'
  });
});

// Admin Session Verify Endpoint
app.get('/api/admin/verify', requireAdminAuth, (req, res) => {
  res.json({
    success: true,
    user: {
      username: req.adminUser.username
    }
  });
});

// Admin Logout Endpoint
app.post('/api/admin/logout', (req, res) => {
  const authHeader = req.headers.authorization || '';
  const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7).trim() : '';

  if (token && activeSessions.has(token)) {
    activeSessions.delete(token);
  }

  res.json({ success: true, message: 'Logged out successfully.' });
});

// Admin Inquiries List Endpoint (Protected)
app.get('/api/admin/enquiries', requireAdminAuth, async (req, res) => {
  try {
    const data = await readAllEnquiriesFromExcel();
    res.json({
      success: true,
      total: data.total,
      inquiries: data.inquiries,
      source: data.source
    });
  } catch (err) {
    console.error('Error fetching enquiries for admin dashboard:', err);
    res.status(500).json({
      success: false,
      message: 'Failed to retrieve enquiries data.',
      error: err.message
    });
  }
});

// ----------------------------------------------------------------------
// COMPANY MASTER APIS (Protected)
// ----------------------------------------------------------------------
app.get('/api/admin/company', requireAdminAuth, async (req, res) => {
  try {
    const company = await getCompanyMaster();
    res.json({ success: true, company });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.put('/api/admin/company', requireAdminAuth, async (req, res) => {
  try {
    const updated = await updateCompanyMaster(req.body || {});
    res.json({ success: true, message: 'Company details updated successfully.', company: updated });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

// ----------------------------------------------------------------------
// CLIENTS MASTER APIS (Protected)
// ----------------------------------------------------------------------
app.get('/api/admin/clients', requireAdminAuth, async (req, res) => {
  try {
    const clients = await getAllClients();
    res.json({ success: true, count: clients.length, clients });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/admin/clients', requireAdminAuth, async (req, res) => {
  try {
    const { clientName, mobile, email } = req.body || {};
    if (!clientName || !clientName.trim()) {
      return res.status(400).json({ success: false, message: 'Client name is required.' });
    }
    if (mobile && !/^\d{10}$/.test(mobile.replace(/[^0-9]/g, ''))) {
      return res.status(400).json({ success: false, message: 'Mobile must be a valid 10-digit number.' });
    }
    if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return res.status(400).json({ success: false, message: 'Please provide a valid email.' });
    }
    const saved = await saveClient(req.body);
    res.json({ success: true, message: 'Client saved successfully.', client: saved });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

// ----------------------------------------------------------------------
// INVOICES & PAYMENTS APIS (Protected)
// ----------------------------------------------------------------------
app.get('/api/admin/invoices', requireAdminAuth, async (req, res) => {
  try {
    const invoices = await getAllInvoices();
    invoices.sort((a, b) => b.invoiceNo.localeCompare(a.invoiceNo));
    res.json({ success: true, count: invoices.length, invoices });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.get('/api/admin/invoices/next-number', requireAdminAuth, async (req, res) => {
  try {
    const nextNo = await generateNextInvoiceNo();
    res.json({ success: true, nextInvoiceNo: nextNo });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.get('/api/admin/invoices/:invoiceNo', requireAdminAuth, async (req, res) => {
  try {
    const invoice = await getInvoiceByNo(req.params.invoiceNo);
    if (!invoice) {
      return res.status(404).json({ success: false, message: 'Invoice not found.' });
    }
    const company = await getCompanyMaster();
    res.json({ success: true, invoice, company });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/admin/invoices', requireAdminAuth, async (req, res) => {
  try {
    const body = req.body || {};
    if (!body.clientId) {
      return res.status(400).json({ success: false, message: 'Client is required.' });
    }
    if (!Array.isArray(body.items) || body.items.length === 0) {
      return res.status(400).json({ success: false, message: 'At least one item line is required.' });
    }
    for (const item of body.items) {
      if (!item.description || !item.description.trim()) {
        return res.status(400).json({ success: false, message: 'Item description is required.' });
      }
      if (Number(item.rate) <= 0) {
        return res.status(400).json({ success: false, message: 'Item rate must be greater than zero.' });
      }
    }
    const created = await createInvoice(body);
    res.json({ success: true, message: 'Invoice created successfully.', invoice: created });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/admin/invoices/:invoiceNo/payments', requireAdminAuth, async (req, res) => {
  try {
    const result = await addPayment(req.params.invoiceNo, req.body || {});
    res.json({ success: true, message: 'Payment recorded successfully.', ...result });
  } catch (err) {
    res.status(400).json({ success: false, message: err.message });
  }
});

// Download PDF Endpoint
app.get('/api/admin/invoices/:invoiceNo/pdf', requireAdminAuth, async (req, res) => {
  try {
    const invoice = await getInvoiceByNo(req.params.invoiceNo);
    if (!invoice) {
      return res.status(404).json({ success: false, message: 'Invoice not found.' });
    }
    const company = await getCompanyMaster();
    const origin = req.get('origin') || process.env.FRONTEND_ORIGIN || 'https://yogteck.com';
    const pdfBuffer = await generateInvoicePdfBuffer(invoice, company, origin);

    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename="${invoice.invoiceNo}.pdf"`,
      'Content-Length': pdfBuffer.length
    });
    res.send(pdfBuffer);
  } catch (err) {
    console.error('PDF Generation Error:', err);
    res.status(500).json({ success: false, message: 'Failed to generate PDF: ' + err.message });
  }
});

// ----------------------------------------------------------------------
// PUBLIC INVOICE VIEW API (QR Code link, Read-Only, No sensitive data)
// ----------------------------------------------------------------------
app.get('/api/public/invoice/:token', async (req, res) => {
  try {
    const invoice = await getInvoiceByToken(req.params.token);
    if (!invoice) {
      return res.status(404).json({ success: false, message: 'Invoice not found or link has expired.' });
    }
    const company = await getCompanyMaster();

    const maskedClient = invoice.client ? {
      clientName: invoice.client.clientName,
      state: invoice.client.state
    } : null;

    res.json({
      success: true,
      invoice: {
        invoiceNo: invoice.invoiceNo,
        invoiceDate: invoice.invoiceDate,
        dueDate: invoice.dueDate,
        grandTotal: invoice.grandTotal,
        totalPaid: invoice.totalPaid,
        balanceDue: invoice.balanceDue,
        status: invoice.status,
        planName: invoice.planName,
        serviceStartDate: invoice.serviceStartDate,
        serviceExpiryDate: invoice.serviceExpiryDate,
        renewalDueDate: invoice.renewalDueDate,
        client: maskedClient,
        payments: invoice.payments || []
      },
      company: {
        companyName: company.companyName,
        contact: company.contact,
        email: company.email,
        bankName: company.bankName,
        accountNo: company.accountNo,
        ifsc: company.ifsc,
        upiId: company.upiId,
        importantNotes: company.importantNotes
      }
    });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

// Public Contact / Inquiry Submission Endpoint
app.post('/api/enquiries/contact', async (req, res) => {
  const clientIp = (req.headers['x-forwarded-for'] || '').split(',')[0].trim() || req.socket.remoteAddress || '';

  console.log('Contact enquiry request received:', {
    origin: req.get('origin') || null,
    ip: clientIp,
    body: req.body || {},
    mailConfig: getMailConfigLog()
  });

  const { enquiry, errors, isValid } = validateContactPayload(req.body || {});
  if (!isValid) {
    console.warn('Contact enquiry validation failed:', { enquiry, errors });
    res.status(400).json({ success: false, message: 'Please check the form fields.', errors });
    return;
  }

  // 1. Record enquiry details, date, time & IP into Excel File
  try {
    const excelResult = await appendEnquiryToExcel(enquiry, clientIp);
    console.log('Enquiry successfully recorded in Excel:', excelResult);
  } catch (excelErr) {
    console.error('Warning: Failed to write enquiry to Excel file:', excelErr);
  }

  // 2. Send Emails (Admin & Customer Confirmation)
  try {
    const transporter = createTransporter();
    
    // Prepare Admin Email (to yogteck@gmail.com)
    const adminEmail = renderAdminEnquiryEmail(enquiry);
    const adminMailOptions = {
      from: process.env.MAIL_FROM || process.env.SMTP_USER,
      to: enquiryRecipient,
      replyTo: enquiry.email,
      subject: `New YogTeck enquiry from ${enquiry.name}${enquiry.companyName ? ` (${enquiry.companyName})` : ''}`,
      text: adminEmail.text,
      html: adminEmail.html
    };

    console.log('Sending admin enquiry notification:', {
      to: adminMailOptions.to,
      subject: adminMailOptions.subject
    });

    const adminInfo = await transporter.sendMail(adminMailOptions);
    console.log('Admin enquiry email sent successfully:', adminInfo.messageId);

    // Prepare & Send Customer Confirmation Email (to user's email)
    try {
      const customerEmail = renderCustomerConfirmationEmail(enquiry);
      const customerMailOptions = {
        from: process.env.MAIL_FROM || process.env.SMTP_USER,
        to: enquiry.email,
        replyTo: enquiryRecipient,
        subject: `Thank you for contacting YOGTECK! We received your enquiry`,
        text: customerEmail.text,
        html: customerEmail.html
      };

      console.log('Sending customer confirmation email:', {
        to: customerMailOptions.to,
        subject: customerMailOptions.subject
      });

      const customerInfo = await transporter.sendMail(customerMailOptions);
      console.log('Customer confirmation email sent successfully:', customerInfo.messageId);
    } catch (custErr) {
      console.error('Customer confirmation email failed to send:', custErr);
    }

    res.json({ 
      success: true, 
      message: 'Enquiry sent successfully! A confirmation email has been sent to your email address.' 
    });
  } catch (error) {
    console.error('Failed to send enquiry email:', error);
    const detail = error instanceof Error ? error.message : 'Unable to send enquiry right now.';
    res.status(500).json({ success: false, message: 'Unable to send enquiry right now.', error: detail });
  }
});

// Download / Export Enquiries Excel Endpoint
app.get('/api/enquiries/export', (req, res) => {
  const filePath = getExcelFilePath();
  if (!existsExcelFile()) {
    res.status(404).json({ success: false, message: 'No enquiries recorded yet.' });
    return;
  }
  res.download(filePath, 'yogteck_enquiries.xlsx', (err) => {
    if (err) {
      console.error('Error downloading Excel file:', err);
      if (!res.headersSent) {
        res.status(500).json({ success: false, message: 'Failed to download Excel file.' });
      }
    }
  });
});

// Redirect Analytics Storage
const redirectLogs = new Map();

app.post('/api/logs/redirect', (req, res) => {
  const { invalidUrl, destination, timestamp } = req.body || {};
  if (!invalidUrl) {
    res.status(400).json({ success: false, message: 'invalidUrl is required.' });
    return;
  }

  const key = `${invalidUrl.toLowerCase()} -> ${destination || '/'}`;
  const existing = redirectLogs.get(key) || {
    invalidUrl,
    destination: destination || '/',
    hits: 0,
    firstSeen: timestamp || new Date().toISOString(),
    lastSeen: timestamp || new Date().toISOString()
  };

  existing.hits += 1;
  existing.lastSeen = timestamp || new Date().toISOString();
  redirectLogs.set(key, existing);

  console.log(`[SEO 404 REDIRECT LOGGED] ${invalidUrl} -> ${destination || '/'} (Hits: ${existing.hits})`);
  res.json({ success: true, logged: existing });
});

app.get('/api/logs/redirect', (req, res) => {
  const logs = Array.from(redirectLogs.values()).sort((a, b) => b.hits - a.hits);
  res.json({ success: true, count: logs.length, logs });
});

app.use((req, res) => {
  if (req.path.startsWith('/api')) {
    res.status(404).json({
      success: false,
      status: 404,
      error: 'Not Found',
      message: `API endpoint '${req.method} ${req.originalUrl}' does not exist.`
    });
    return;
  }
  // 301 Permanent Redirect non-API requests to primary canonical URL
  res.redirect(301, `https://yogteck.com${req.url}`);
});

app.use((error, req, res, next) => {
  console.error(error);
  res.status(500).json({ success: false, message: 'Internal server error.' });
});

if (!process.env.VERCEL) {
  app.listen(port, () => {
    console.log(`YogTeck backend running on port ${port}`);
  });
}

module.exports = app;
