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
  'https://yogteck-frontend.vercel.app'
];

const allowedOrigins = [
  ...defaultFrontendOrigins,
  ...(process.env.FRONTEND_ORIGIN || '').split(',')
]
  .map(origin => origin.trim())
  .filter(Boolean);

app.use(helmet());
app.use(cors({
  origin(origin, callback) {
    if (!origin || allowedOrigins.includes(origin)) {
      callback(null, true);
      return;
    }
    callback(new Error('Not allowed by CORS'));
  }
}));
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
    phone: cleanString(body.phone),
    email: cleanString(body.email),
    rackType: cleanString(body.rackType),
    message: cleanString(body.message)
  };

  const errors = {};
  if (!enquiry.name) errors.name = 'Name is required.';
  if (!enquiry.phone) errors.phone = 'Mobile number is required.';
  if (!enquiry.email) {
    errors.email = 'Email is required.';
  } else if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(enquiry.email)) {
    errors.email = 'Enter a valid email address.';
  }
  if (!enquiry.rackType) errors.rackType = 'Rack type is required.';
  if (!enquiry.message) errors.message = 'Message is required.';

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

function renderContactEmail(enquiry) {
  const rows = [
    ['Name', enquiry.name],
    ['Mobile Number', enquiry.phone],
    ['Email', enquiry.email],
    ['Rack Type', enquiry.rackType || 'Not selected'],
    ['Message', enquiry.message]
  ];

  const htmlRows = rows.map(([label, value]) => `
    <tr>
      <td style="padding:10px 12px;border:1px solid #ddd;font-weight:700;">${escapeHtml(label)}</td>
      <td style="padding:10px 12px;border:1px solid #ddd;">${escapeHtml(value).replace(/\n/g, '<br>')}</td>
    </tr>
  `).join('');

  const text = rows.map(([label, value]) => `${label}: ${value || 'Not selected'}`).join('\n');

  return {
    text,
    html: `
      <div style="font-family:Arial,sans-serif;line-height:1.5;color:#222;">
        <h2>New YogTeck Website Enquiry</h2>
        <table style="border-collapse:collapse;width:100%;max-width:680px;">${htmlRows}</table>
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
    const email = renderContactEmail(enquiry);

    const mailOptions = {
      from: process.env.MAIL_FROM || process.env.SMTP_USER,
      to: enquiryRecipient,
      replyTo: enquiry.email,
      subject: `New YogTeck enquiry from ${enquiry.name}`,
      text: email.text,
      html: email.html
    };

    console.log('Sending contact enquiry email:', {
      from: mailOptions.from,
      to: mailOptions.to,
      replyTo: mailOptions.replyTo,
      subject: mailOptions.subject,
      enquiry
    });

    const info = await transporter.sendMail(mailOptions);

    console.log('Contact enquiry email sent:', {
      messageId: info.messageId,
      accepted: info.accepted,
      rejected: info.rejected,
      response: info.response
    });

    res.json({ success: true, message: 'Enquiry sent successfully.' });
  } catch (error) {
    console.error('Failed to send enquiry email:', error);
    res.status(500).json({ success: false, message: 'Unable to send enquiry right now.' });
  }
});

app.use((req, res) => {
  res.status(404).json({ success: false, message: 'API route not found.' });
});

app.use((error, req, res, next) => {
  console.error(error);
  res.status(500).json({ success: false, message: 'Internal server error.' });
});

app.listen(port, () => {
  console.log(`YogTeck backend running on port ${port}`);
});
