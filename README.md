# YogTeck Backend

Backend APIs for YogTeck website forms. The contact/enquiry form submits to this backend and the enquiry is sent to email.

## Setup

```bash
npm install
copy .env.example .env
npm run dev
```

Update `.env` with real SMTP details before using the API.

For Gmail, use an App Password in `SMTP_PASS`; normal account passwords usually do not work.

## API

### Health

`GET /api/health`

### Contact enquiry

`POST /api/enquiries/contact`

Body:

```json
{
  "name": "Customer Name",
  "phone": "+91 99999 99999",
  "email": "customer@example.com",
  "rackType": "Wall Display Racks",
  "message": "I need racks for my store."
}
```

Success response:

```json
{
  "success": true,
  "message": "Enquiry sent successfully."
}
```
