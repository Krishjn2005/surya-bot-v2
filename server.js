const express = require('express');
const twilio = require('twilio');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const Anthropic = require('@anthropic-ai/sdk');
require('dotenv').config();

const app = express();
app.use(express.urlencoded({ extended: false }));

// Root redirect
app.get('/', (req, res) => {
  res.redirect('/dashboard');
});

const client = twilio(process.env.TWILIO_ACCOUNT_SID, process.env.TWILIO_AUTH_TOKEN);
const anthropic = new Anthropic({
  apiKey: process.env.CLAUDE_API_KEY,
});

const TWILIO_WHATSAPP_NUMBER = process.env.TWILIO_WHATSAPP_NUMBER;
const MANAGER_PHONE = process.env.MANAGER_PHONE;
const SHEET_ID = process.env.GOOGLE_SHEET_ID;
const GOOGLE_SERVICE_ACCOUNT = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT || '{}');

let doc;
let registrationSheet, approvedSheet, complaintsSheet;

// Initialize Google Sheets
async function initializeSheets() {
  try {
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
  } catch (error) {
    console.error('Error initializing sheets:', error);
  }
}

// Temporarily disabled - uncomment when Google Sheets is properly configured
// initializeSheets();

// Process incoming WhatsApp messages
app.post('/whatsapp', async (req, res) => {
  const from = req.body.From.replace('whatsapp:', '');
  const messageBody = req.body.Body.trim();

  try {
    // Check if user is approved
    const approvedRows = approvedSheet ? await approvedSheet.getRows() : [];
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
  const nameMatch = messageBody.match(/name\s+(?:is\s+)?([^,]+)/i);
  const flatMatch = messageBody.match(/flat\s+([^,]+)/i);
  const phoneMatch = messageBody.match(/phone\s+(\d+)/);

  if (nameMatch && flatMatch && phoneMatch) {
    const name = nameMatch[1].trim();
    const flat = flatMatch[1].trim();
    const phone = phoneMatch[1];

    // Add to registration sheet (pending approval)
    if (registrationSheet) {
      await registrationSheet.addRow({
        Phone: from,
        Name: name,
        Flat: flat,
        Status: 'Pending',
        Timestamp: new Date().toISOString(),
      });
    }

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
  const approvedRows = approvedSheet ? await approvedSheet.getRows() : [];
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
  if (complaintsSheet) {
    await complaintsSheet.addRow({
      Flat: resident.Flat,
      Phone: from,
      Name: resident.Name,
      'Issue Type': complaintType,
      Description: description,
      Status: 'Pending',
      Timestamp: new Date().toISOString(),
    });
  }

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
  try {
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
  } catch (error) {
    console.error('Claude API error:', error);
    return message;
  }
}

// Send WhatsApp message
async function sendWhatsAppMessage(to, message) {
  try {
    await client.messages.create({
      body: message,
      from: `whatsapp:${TWILIO_WHATSAPP_NUMBER}`,
      to: `whatsapp:${to}`,
    });
  } catch (error) {
    console.error('Error sending message:', error);
  }
}

// Beautiful Modern Dashboard
app.get('/dashboard', async (req, res) => {
  const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Surya Prakash Residency - Manager Dashboard</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); min-height: 100vh; padding: 20px; }
    .container { max-width: 1400px; margin: 0 auto; }
    .header { background: white; padding: 30px; border-radius: 12px; margin-bottom: 30px; box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1); }
    .header h1 { color: #1a1a2e; font-size: 28px; margin-bottom: 5px; }
    .header p { color: #666; font-size: 14px; }
    .stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-top: 20px; }
    .stat-card { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; border-radius: 8px; text-align: center; }
    .stat-card h3 { font-size: 12px; opacity: 0.9; margin-bottom: 8px; text-transform: uppercase; letter-spacing: 1px; }
    .stat-card .number { font-size: 32px; font-weight: bold; }
    .section { background: white; padding: 30px; border-radius: 12px; margin-bottom: 30px; box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1); }
    .section h2 { color: #1a1a2e; font-size: 22px; margin-bottom: 25px; display: flex; align-items: center; gap: 10px; }
    .section h2::before { content: ''; display: inline-block; width: 4px; height: 24px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 2px; }
    table { width: 100%; border-collapse: collapse; }
    thead { background: #f8f9fa; }
    th { padding: 15px; text-align: left; font-weight: 600; color: #1a1a2e; font-size: 14px; border-bottom: 2px solid #e9ecef; }
    td { padding: 15px; border-bottom: 1px solid #e9ecef; color: #333; }
    tr:hover { background: #f8f9fa; }
    .btn { padding: 8px 16px; border: none; border-radius: 6px; font-size: 13px; font-weight: 600; cursor: pointer; transition: all 0.3s ease; }
    .btn-approve { background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%); color: white; margin-right: 8px; }
    .btn-approve:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(17, 153, 142, 0.3); }
    .btn-resolve { background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%); color: white; }
    .empty-state { text-align: center; padding: 60px 20px; color: #999; }
    .modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0, 0, 0, 0.5); justify-content: center; align-items: center; z-index: 1000; }
    .modal.active { display: flex; }
    .modal-content { background: white; padding: 30px; border-radius: 12px; max-width: 400px; text-align: center; box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3); }
    .modal-buttons { display: flex; gap: 10px; margin-top: 20px; justify-content: center; }
    .modal-buttons button { flex: 1; padding: 10px; border: none; border-radius: 6px; cursor: pointer; font-weight: 600; }
    .modal-buttons .confirm { background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%); color: white; }
    .modal-buttons .cancel { background: #eee; color: #333; }
    @media (max-width: 768px) { .section { padding: 20px; } th, td { padding: 10px; font-size: 13px; } .header h1 { font-size: 22px; } }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>🏢 Surya Prakash Residency</h1>
      <p>Manager Dashboard • WhatsApp Complaint Management System</p>
      <div class="stats">
        <div class="stat-card"><h3>Pending Approvals</h3><div class="number">0</div></div>
        <div class="stat-card"><h3>Active Complaints</h3><div class="number">0</div></div>
        <div class="stat-card"><h3>Approved Residents</h3><div class="number">0</div></div>
        <div class="stat-card"><h3>Resolved</h3><div class="number">0</div></div>
      </div>
    </div>

    <div class="section">
      <h2>📋 Pending Registrations</h2>
      <table>
        <thead>
          <tr><th>Name</th><th>Flat</th><th>Phone</th><th>Date</th><th>Action</th></tr>
        </thead>
        <tbody>
          <tr><td colspan="5" class="empty-state"><h3>✨ No pending registrations</h3><p>New registration requests will appear here</p></td></tr>
        </tbody>
      </table>
    </div>

    <div class="section">
      <h2>⚠️ Active Complaints</h2>
      <table>
        <thead>
          <tr><th>Flat</th><th>Resident</th><th>Issue</th><th>Type</th><th>Date</th><th>Action</th></tr>
        </thead>
        <tbody>
          <tr><td colspan="6" class="empty-state"><h3>✨ No active complaints</h3><p>Resident complaints will appear here as they submit them</p></td></tr>
        </tbody>
      </table>
    </div>
  </div>
</body>
</html>`;
  res.send(html);
});

// API to approve resident
app.post('/api/approve', async (req, res) => {
  const { flat, phone } = req.body;

  try {
    if (!registrationSheet || !approvedSheet) {
      return res.status(500).send('Database not initialized');
    }

    const registrations = await registrationSheet.getRows();
    const registration = registrations.find(r => r.Flat === flat);

    if (registration) {
      await approvedSheet.addRow({
        Phone: phone,
        Name: registration.Name,
        Flat: flat,
        'Approved Date': new Date().toISOString(),
      });

      registration.Status = 'Approved';
      await registration.save();

      await sendWhatsAppMessage(
        phone,
        `✅ Great news, Flat ${flat}! Your registration has been approved. You can now file complaints and inquiries. 🎉\n\nJust text your issue (e.g., "Gym AC kharab hai", "Plumber chahiye") and we'll help!`
      );
    }

    res.status(200).send('OK');
  } catch (error) {
    console.error('Error approving resident:', error);
    res.status(500).send('Error');
  }
});

// API to resolve complaint
app.post('/api/resolve', async (req, res) => {
  const { flat } = req.body;

  try {
    if (!complaintsSheet) {
      return res.status(500).send('Database not initialized');
    }

    const complaints = await complaintsSheet.getRows();
    const complaint = complaints.find(c => c.Flat === flat && c.Status !== 'Resolved');

    if (complaint) {
      complaint.Status = 'Resolved';
      await complaint.save();

      await sendWhatsAppMessage(
        complaint.Phone,
        `✅ Your complaint for Flat ${flat} has been resolved! Thank you for reporting. 🙏`
      );
    }

    res.status(200).send('OK');
  } catch (error) {
    console.error('Error resolving complaint:', error);
    res.status(500).send('Error');
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));