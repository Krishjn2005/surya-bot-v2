const express = require('express');
const twilio = require('twilio');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const Anthropic = require('@anthropic-ai/sdk');
require('dotenv').config();

const app = express();
app.use(express.urlencoded({ extended: false }));

const client = twilio.Twilio(process.env.TWILIO_ACCOUNT_SID, process.env.TWILIO_AUTH_TOKEN);
const anthropic = new Anthropic({
  apiKey: process.env.CLAUDE_API_KEY,
});

const TWILIO_WHATSAPP_NUMBER = process.env.TWILIO_WHATSAPP_NUMBER;
const MANAGER_PHONE = process.env.MANAGER_PHONE;
const SHEET_ID = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SERVICE_ACCOUNT = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);

let doc;
let registrationSheet, approvedSheet, complaintsSheet;

// Initialize Google Sheets
async function initializeSheets() {
  doc = new GoogleSpreadsheet(SHEET_ID);
  await doc.useServiceAccountAuth(GOOGLE_SERVICE_ACCOUNT);
  await doc.loadInfo();

  registrationSheet = doc.sheetsByTitle['Registrations'] || await doc.addSheet({ title: 'Registrations' });
  approvedSheet = doc.sheetsByTitle['Approved Residents'] || await doc.addSheet({ title: 'Approved Residents' });
  complaintsSheet = doc.sheetsByTitle['Complaints'] || await doc.addSheet({ title: 'Complaints' });

  // Add headers if empty
  if (registrationSheet.rowCount === 1 && !registrationSheet.headerValues.length) {
    await registrationSheet.setHeaderRow(['Phone', 'Name', 'Flat', 'Status', 'Timestamp']);
  }
  if (approvedSheet.rowCount === 1 && !approvedSheet.headerValues.length) {
    await approvedSheet.setHeaderRow(['Phone', 'Name', 'Flat', 'Approved Date']);
  }
  if (complaintsSheet.rowCount === 1 && !complaintsSheet.headerValues.length) {
    await complaintsSheet.setHeaderRow(['Flat', 'Phone', 'Name', 'Issue Type', 'Description', 'Status', 'Timestamp']);
  }
}

initializeSheets();

// Process incoming WhatsApp messages
app.post('/whatsapp', async (req, res) => {
  const from = req.body.From.replace('whatsapp:', '');
  const messageBody = req.body.Body.trim();

  try {
    // Check if user is approved
    const approvedRows = await approvedSheet.getRows();
    const isApproved = approvedRows.some(row => row.Phone === from);

    if (isApproved) {
      // Handle complaint filing
      await handleComplaintFiling(from, messageBody);
    } else {
      // Handle registration
      await handleRegistration(from, messageBody);
    }

    res.status(200).send('OK');
  } catch (error) {
    console.error('Error processing message:', error);
    res.status(500).send('Error');
  }
});

// Registration flow
async function handleRegistration(from, messageBody) {
  // Parse: "My name is Amit, flat 405, phone 9876543210"
  const nameMatch = messageBody.match(/name\s+(?:is\s+)?([^,]+)/i);
  const flatMatch = messageBody.match(/flat\s+([^,]+)/i);
  const phoneMatch = messageBody.match(/phone\s+(\d+)/);

  if (nameMatch && flatMatch && phoneMatch) {
    const name = nameMatch[1].trim();
    const flat = flatMatch[1].trim();
    const phone = phoneMatch[1];

    // Add to registration sheet (pending approval)
    await registrationSheet.addRow({
      Phone: from,
      Name: name,
      Flat: flat,
      Status: 'Pending',
      Timestamp: new Date().toISOString(),
    });

    // Reply to resident
    await sendWhatsAppMessage(
      from,
      `Hi ${name}! Your registration request for Flat ${flat} has been sent to the manager. You'll get access within 24 hours. 🎯`
    );

    // Alert manager
    await sendWhatsAppMessage(
      MANAGER_PHONE,
      `📋 NEW REGISTRATION REQUEST\n\nName: ${name}\nFlat: ${flat}\nPhone: ${phone}\n\nReply with "approve ${flat}" or "deny ${flat}" to manage access.`
    );
  } else {
    await sendWhatsAppMessage(
      from,
      `Hi! To register, please send: "My name is [Your Name], flat [Number], phone [Phone]"\n\nExample: "My name is Amit, flat 405, phone 9876543210"`
    );
  }
}

// Complaint filing for approved residents
async function handleComplaintFiling(from, messageBody) {
  // Get resident info
  const approvedRows = await approvedSheet.getRows();
  const resident = approvedRows.find(row => row.Phone === from);

  if (!resident) {
    await sendWhatsAppMessage(from, 'Access denied. Please contact the society office.');
    return;
  }

  // Use Claude to understand complaint in Hinglish
  const description = await processComplaintWithClaude(messageBody);

  // Determine complaint type
  let complaintType = 'Other';
  const lowerMsg = messageBody.toLowerCase();
  if (lowerMsg.includes('gym') || lowerMsg.includes('electrician') || lowerMsg.includes('plumber')) {
    complaintType = 'Maintenance';
  }
  if (lowerMsg.includes('clubhouse') || lowerMsg.includes('parking') || lowerMsg.includes('gym')) {
    complaintType = 'Amenity';
  }

  // Log complaint
  await complaintsSheet.addRow({
    Flat: resident.Flat,
    Phone: from,
    Name: resident.Name,
    'Issue Type': complaintType,
    Description: description,
    Status: 'Pending',
    Timestamp: new Date().toISOString(),
  });

  // Reply to resident (in Hinglish)
  await sendWhatsAppMessage(
    from,
    `✅ Complaint received for Flat ${resident.Flat}!\n\n"${description}"\n\nManager ko bhej diya. 24 ghanton mein reply milega. 🔔`
  );

  // Alert manager
  await sendWhatsAppMessage(
    MANAGER_PHONE,
    `⚠️ NEW COMPLAINT\n\nFlat: ${resident.Flat}\nResident: ${resident.Name}\nIssue: ${description}\n\nReply "resolve ${resident.Flat}" when done.`
  );
}

// Use Claude to understand Hinglish complaints
async function processComplaintWithClaude(message) {
  const response = await anthropic.messages.create({
    model: 'claude-opus-4-1',
    max_tokens: 100,
    messages: [
      {
        role: 'user',
        content: `You are a complaint analyzer for a society. Understand this Hinglish/Hindi complaint and summarize it clearly in English. Keep it short (one line). Return ONLY the summary, nothing else.\n\nComplaint: "${message}"`,
      },
    ],
  });

  return response.content[0].type === 'text' ? response.content[0].text : message;
}

// Send WhatsApp message
async function sendWhatsAppMessage(to, message) {
  await client.messages.create({
    body: message,
    from: `whatsapp:${TWILIO_WHATSAPP_NUMBER}`,
    to: `whatsapp:${to}`,
  });
}

// Manager dashboard endpoint
app.get('/dashboard', async (req, res) => {
  const registrations = await registrationSheet.getRows();
  const complaints = await complaintsSheet.getRows();

  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <title>Surya Prakash Residency - Manager Dashboard</title>
      <style>
        body { font-family: Arial; margin: 20px; background: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; }
        h1 { color: #333; }
        .section { background: white; padding: 20px; margin: 20px 0; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background: #1F4788; color: white; }
        tr:hover { background: #f9f9f9; }
        .pending { color: orange; font-weight: bold; }
        .approved { color: green; font-weight: bold; }
        .button { padding: 8px 16px; background: #1F4788; color: white; border: none; border-radius: 4px; cursor: pointer; }
        .button:hover { background: #153060; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>🏢 Surya Prakash Residency - Manager Dashboard</h1>

        <div class="section">
          <h2>📋 Pending Registrations</h2>
          <table>
            <tr>
              <th>Name</th>
              <th>Flat</th>
              <th>Phone</th>
              <th>Status</th>
              <th>Action</th>
            </tr>
            ${registrations.filter(r => r.Status === 'Pending').map(r => `
              <tr>
                <td>${r.Name}</td>
                <td>${r.Flat}</td>
                <td>${r.Phone}</td>
                <td><span class="pending">${r.Status}</span></td>
                <td>
                  <button class="button" onclick="approveResident('${r.Flat}', '${r.Phone}')">Approve</button>
                </td>
              </tr>
            `).join('')}
          </table>
        </div>

        <div class="section">
          <h2>⚠️ Active Complaints</h2>
          <table>
            <tr>
              <th>Flat</th>
              <th>Resident</th>
              <th>Issue</th>
              <th>Type</th>
              <th>Status</th>
              <th>Date</th>
              <th>Action</th>
            </tr>
            ${complaints.filter(c => c.Status !== 'Resolved').map(c => `
              <tr>
                <td>${c.Flat}</td>
                <td>${c.Name}</td>
                <td>${c.Description}</td>
                <td>${c['Issue Type']}</td>
                <td>${c.Status}</td>
                <td>${new Date(c.Timestamp).toLocaleDateString()}</td>
                <td>
                  <button class="button" onclick="resolveComplaint('${c.Flat}')">Resolve</button>
                </td>
              </tr>
            `).join('')}
          </table>
        </div>
      </div>

      <script>
        function approveResident(flat, phone) {
          fetch('/api/approve', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ flat, phone })
          }).then(() => location.reload());
        }

        function resolveComplaint(flat) {
          fetch('/api/resolve', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ flat })
          }).then(() => location.reload());
        }
      </script>
    </body>
    </html>
  `;

  res.send(html);
});

// API to approve resident
app.post('/api/approve', async (req, res) => {
  const { flat, phone } = req.body;

  // Move from registration to approved
  const registrations = await registrationSheet.getRows();
  const registration = registrations.find(r => r.Flat === flat);

  if (registration) {
    await approvedSheet.addRow({
      Phone: phone,
      Name: registration.Name,
      Flat: flat,
      'Approved Date': new Date().toISOString(),
    });

    // Remove from registration sheet (update status instead)
    registration.Status = 'Approved';
    await registration.save();

    // Notify resident
    await sendWhatsAppMessage(
      phone,
      `✅ Great news, Flat ${flat}! Your registration has been approved. You can now file complaints and inquiries. 🎉\n\nJust text your issue (e.g., "Gym AC kharab hai", "Plumber chahiye") and we'll help!`
    );
  }

  res.status(200).send('OK');
});

// API to resolve complaint
app.post('/api/resolve', async (req, res) => {
  const { flat } = req.body;

  const complaints = await complaintsSheet.getRows();
  const complaint = complaints.find(c => c.Flat === flat && c.Status !== 'Resolved');

  if (complaint) {
    complaint.Status = 'Resolved';
    await complaint.save();

    // Notify resident
    await sendWhatsAppMessage(
      complaint.Phone,
      `✅ Your complaint for Flat ${flat} has been resolved! Thank you for reporting. 🙏`
    );
  }

  res.status(200).send('OK');
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
