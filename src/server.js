require('dotenv').config();

const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const nodemailer = require('nodemailer');

const app = express();
const port = Number(process.env.PORT || 5000);
const enquiryRecipient = process.env.MAIL_TO || 'yogteck@gmail.com';

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

// Allow CORS with dynamic origin reflection for public enquiry API
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
    text: `Dear ${enquiry.name},\n\nThank you for reaching out to YOGTECK! We have received your inquiry regarding "${enquiry.rackType || 'Website & Digital Solutions'}".\n\nOur solutions team is reviewing your requirements and will contact you within 2-4 business hours.\n\nSummary of your request:\n- Required Service: ${enquiry.rackType}\n- Contact Phone: ${enquiry.phone}\n${enquiry.companyName ? `- Company: ${enquiry.companyName}\n` : ''}\nIf you have urgent questions, feel free to WhatsApp us directly at +91 8299209905.\n\nBest regards,\nYOGTECK Digital Solutions Team\nhttps://yogteck.com\nEmail: yogteck@gmail.com\nPhone: +91 8299209905\nAddress: H-14, Sector 63, Noida, UP - 201301`,
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
          <div><strong>Address:</strong> H-14, Sector 63, Noida, Uttar Pradesh - 201301</div>
          <div><strong>Phone:</strong> +91 8299209905 &bull; <strong>Email:</strong> yogteck@gmail.com</div>
          <div><strong>Website:</strong> <a href="https://yogteck.com" style="color:#ff9a1f;text-decoration:none;">https://yogteck.com</a></div>
        </div>

      </div>
    `
  };
}

app.get('/api/health', (req, res) => {
  res.json({ success: true, message: 'YogTeck backend is running.' });
});

app.post('/api/enquiries/contact', async (req, res) => {
  console.log('Contact enquiry request received:', {
    origin: req.get('origin') || null,
    body: req.body || {},
    mailConfig: getMailConfigLog()
  });

  const { enquiry, errors, isValid } = validateContactPayload(req.body || {});
  if (!isValid) {
    console.warn('Contact enquiry validation failed:', { enquiry, errors });
    res.status(400).json({ success: false, message: 'Please check the form fields.', errors });
    return;
  }

  try {
    const transporter = createTransporter();
    
    // 1. Prepare Admin Email (to yogteck@gmail.com)
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

    // 2. Prepare & Send Customer Confirmation Email (to user's email)
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
    res.status(404).json({ success: false, message: 'API route not found.' });
    return;
  }
  // 301 Permanent Redirect non-API requests to primary canonical URL
  res.redirect(301, `https://yogteck.com${req.url}`);
});

app.use((error, req, res, next) => {
  console.error(error);
  res.status(500).json({ success: false, message: 'Internal server error.' });
});

app.listen(port, () => {
  console.log(`YogTeck backend running on port ${port}`);
});
