const express = require('express');
const twilio = require('twilio');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const Anthropic = require('@anthropic-ai/sdk');
require('dotenv').config();

const app = express();
app.use(express.urlencoded({ extended: false }));

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
let registrationSheet, approvedSheet, complaintsSheet, announcementsSheet;

async function initializeSheets() {
  try {
    doc = new GoogleSpreadsheet(SHEET_ID);
    await doc.useServiceAccountAuth(GOOGLE_SERVICE_ACCOUNT);
    await doc.loadInfo();

    registrationSheet = doc.sheetsByTitle['Registrations'] || await doc.addSheet({ title: 'Registrations' });
    approvedSheet = doc.sheetsByTitle['Approved Residents'] || await doc.addSheet({ title: 'Approved Residents' });
    complaintsSheet = doc.sheetsByTitle['Complaints'] || await doc.addSheet({ title: 'Complaints' });
    announcementsSheet = doc.sheetsByTitle['Announcements'] || await doc.addSheet({ title: 'Announcements' });

    if (registrationSheet.rowCount === 1 && !registrationSheet.headerValues.length) {
      await registrationSheet.setHeaderRow(['Phone', 'Name', 'Flat', 'Status', 'Timestamp']);
    }
    if (approvedSheet.rowCount === 1 && !approvedSheet.headerValues.length) {
      await approvedSheet.setHeaderRow(['Phone', 'Name', 'Flat', 'Approved Date']);
    }
    if (complaintsSheet.rowCount === 1 && !complaintsSheet.headerValues.length) {
      await complaintsSheet.setHeaderRow(['Flat', 'Phone', 'Name', 'Issue Type', 'Description', 'Status', 'Timestamp', 'Priority']);
    }
    if (announcementsSheet.rowCount === 1 && !announcementsSheet.headerValues.length) {
      await announcementsSheet.setHeaderRow(['Title', 'Message', 'Created Date', 'Sent to All', 'Status']);
    }

    console.log('Google Sheets initialized successfully');
  } catch (error) {
    console.error('Error initializing sheets:', error.message);
  }
}

// initializeSheets();

// WhatsApp Message Handler
app.post('/whatsapp', async (req, res) => {
  const from = req.body.From.replace('whatsapp:', '');
  const messageBody = req.body.Body.trim().toLowerCase();

  try {
    const approvedRows = await approvedSheet.getRows();
    const isApproved = approvedRows.some(row => row.Phone === from);

    // Emergency Keywords
    if (messageBody.includes('emergency') || messageBody.includes('urgent') || messageBody.includes('help')) {
      await handleEmergency(from, messageBody);
    } else if (isApproved) {
      await handleComplaintFiling(from, messageBody);
    } else {
      await handleRegistration(from, messageBody);
    }

    res.status(200).send('OK');
  } catch (error) {
    console.error('Error processing message:', error);
    res.status(500).send('Error');
  }
});

// Emergency Contact Handler
async function handleEmergency(from, messageBody) {
  const approvedRows = await approvedSheet.getRows();
  const resident = approvedRows.find(row => row.Phone === from);

  const residentName = resident ? resident.Name : 'Unknown';
  const residentFlat = resident ? resident.Flat : 'Unknown';

  // Log as priority complaint
  await complaintsSheet.addRow({
    Flat: residentFlat,
    Phone: from,
    Name: residentName,
    'Issue Type': 'Emergency',
    Description: messageBody,
    Status: 'Urgent',
    Timestamp: new Date().toISOString(),
    Priority: 'HIGH',
  });

  // Immediate alert to manager
  await sendWhatsAppMessage(
    MANAGER_PHONE,
    `EMERGENCY ALERT\n\nFlat: ${residentFlat}\nResident: ${residentName}\nMessage: ${messageBody}\n\nImmediate action required!`
  );

  // Confirm to resident
  await sendWhatsAppMessage(
    from,
    `Emergency alert sent to manager. Help is on the way!`
  );
}

// Registration Handler
async function handleRegistration(from, messageBody) {
  const nameMatch = messageBody.match(/name\s+(?:is\s+)?([^,]+)/i);
  const flatMatch = messageBody.match(/flat\s+([^,]+)/i);
  const phoneMatch = messageBody.match(/phone\s+(\d+)/);

  if (nameMatch && flatMatch && phoneMatch) {
    const name = nameMatch[1].trim();
    const flat = flatMatch[1].trim();
    const phone = phoneMatch[1];

    await registrationSheet.addRow({
      Phone: from,
      Name: name,
      Flat: flat,
      Status: 'Pending',
      Timestamp: new Date().toISOString(),
    });

    await sendWhatsAppMessage(
      from,
      `Hello ${name}! Your registration request for Flat ${flat} has been submitted. The manager will review and approve your access within 24 hours.`
    );

    await sendWhatsAppMessage(
      MANAGER_PHONE,
      `New Registration Request\n\nName: ${name}\nFlat: ${flat}\nPhone: ${phone}\n\nPlease review this request in the dashboard.`
    );
  } else {
    await sendWhatsAppMessage(
      from,
      `Please send your registration in this format:\nMy name is [Your Name], flat [Number], phone [Phone]\n\nExample: My name is Amit, flat 405, phone 9876543210`
    );
  }
}

// Complaint Filing Handler
async function handleComplaintFiling(from, messageBody) {
  const approvedRows = await approvedSheet.getRows();
  const resident = approvedRows.find(row => row.Phone === from);

  if (!resident) {
    await sendWhatsAppMessage(from, 'Access denied. Please contact the society office.');
    return;
  }

  const description = await processComplaintWithClaude(messageBody);

  let complaintType = 'Other';
  const lowerMsg = messageBody.toLowerCase();
  if (lowerMsg.includes('gym') || lowerMsg.includes('electrician') || lowerMsg.includes('plumber')) {
    complaintType = 'Maintenance';
  }
  if (lowerMsg.includes('clubhouse') || lowerMsg.includes('parking')) {
    complaintType = 'Amenity';
  }

  await complaintsSheet.addRow({
    Flat: resident.Flat,
    Phone: from,
    Name: resident.Name,
    'Issue Type': complaintType,
    Description: description,
    Status: 'Pending',
    Timestamp: new Date().toISOString(),
    Priority: 'Normal',
  });

  await sendWhatsAppMessage(
    from,
    `Complaint received for Flat ${resident.Flat}: "${description}"\n\nYour complaint has been forwarded to the manager. You will receive an update within 24 hours.`
  );

  await sendWhatsAppMessage(
    MANAGER_PHONE,
    `New Complaint\n\nFlat: ${resident.Flat}\nResident: ${resident.Name}\nIssue: ${description}\nType: ${complaintType}\n\nPlease review in the dashboard.`
  );
}

// Claude Hinglish Processing
async function processComplaintWithClaude(message) {
  try {
    const response = await anthropic.messages.create({
      model: 'claude-opus-4-1',
      max_tokens: 100,
      messages: [
        {
          role: 'user',
          content: `Understand this complaint in Hindi or English and summarize it clearly in one sentence. Return only the summary.\n\nComplaint: "${message}"`,
        },
      ],
    });

    return response.content[0].type === 'text' ? response.content[0].text : message;
  } catch (error) {
    console.error('Claude API error:', error);
    return message;
  }
}

// Send WhatsApp Message
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

// Voice Call Handler (Retell Integration)
app.post('/voice-webhook', async (req, res) => {
  const from = req.body.From;
  const callSid = req.body.CallSid;

  try {
    const approvedRows = await approvedSheet.getRows();
    const isApproved = approvedRows.some(row => row.Phone === from.replace('whatsapp:', ''));

    if (isApproved) {
      // Connected to Retell AI agent for complaint filing
      await sendWhatsAppMessage(
        from,
        `Your call has been connected to our complaint handling system. Please describe your issue.`
      );
    } else {
      // New resident - need registration
      await sendWhatsAppMessage(
        from,
        `Please register first via WhatsApp: My name is [Name], flat [Number], phone [Phone]`
      );
    }

    res.status(200).send('OK');
  } catch (error) {
    console.error('Voice webhook error:', error);
    res.status(500).send('Error');
  }
});

// Broadcast Announcement Endpoint
app.post('/api/broadcast-announcement', async (req, res) => {
  const { title, message } = req.body;

  try {
    const approvedRows = await approvedSheet.getRows();

    // Log announcement
    await announcementsSheet.addRow({
      Title: title,
      Message: message,
      'Created Date': new Date().toISOString(),
      'Sent to All': approvedRows.length,
      Status: 'Sent',
    });

    // Send to all approved residents
    for (const resident of approvedRows) {
      await sendWhatsAppMessage(
        resident.Phone,
        `Announcement:\n\n${title}\n\n${message}`
      );
    }

    // Notify manager
    await sendWhatsAppMessage(
      MANAGER_PHONE,
      `Announcement sent to ${approvedRows.length} residents:\n${title}`
    );

    res.status(200).send('OK');
  } catch (error) {
    console.error('Broadcast error:', error);
    res.status(500).send('Error');
  }
});

// Dashboard with Phase 2 Features
app.get('/dashboard', async (req, res) => {
  const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Surya Prakash Residency - Manager Dashboard</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background-color: #f5f5f5; color: #1a1a1a; }
    .navbar { background-color: #000; color: #fff; padding: 20px 40px; }
    .navbar h1 { font-size: 24px; font-weight: 600; }
    .navbar p { font-size: 13px; color: #aaa; margin-top: 4px; }
    .container { max-width: 1200px; margin: 0 auto; padding: 40px 20px; }
    .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 20px; margin-bottom: 40px; }
    .stat-card { background: #fff; padding: 25px; border-radius: 8px; border-left: 4px solid #000; box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05); }
    .stat-card h3 { font-size: 12px; font-weight: 600; text-transform: uppercase; letter-spacing: 1px; color: #666; margin-bottom: 12px; }
    .stat-card .number { font-size: 36px; font-weight: 700; color: #000; }
    .section { background: #fff; border-radius: 8px; padding: 30px; margin-bottom: 30px; box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05); }
    .section h2 { font-size: 18px; font-weight: 600; margin-bottom: 25px; color: #000; }
    .form-group { margin-bottom: 15px; }
    .form-group label { display: block; font-size: 13px; font-weight: 600; margin-bottom: 6px; }
    .form-group input, .form-group textarea { width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; }
    .form-group textarea { resize: vertical; min-height: 80px; }
    .btn { padding: 10px 20px; background-color: #000; color: #fff; border: none; border-radius: 4px; cursor: pointer; font-weight: 600; }
    .btn:hover { background-color: #333; }
    .btn-emergency { background-color: #c33; }
    .btn-emergency:hover { background-color: #a00; }
    .empty-state { text-align: center; padding: 50px 20px; color: #999; }
    .empty-state h3 { font-size: 16px; font-weight: 500; color: #666; margin-bottom: 8px; }
    .emergency-alert { background-color: #ffe6e6; border-left: 4px solid #c33; padding: 15px; margin-bottom: 15px; border-radius: 4px; }
    .emergency-alert h4 { color: #c33; font-weight: 600; margin-bottom: 5px; }
  </style>
</head>
<body>
  <div class="navbar">
    <h1>Surya Prakash Residency</h1>
    <p>Manager Dashboard - Phase 2 Complete</p>
  </div>
  <div class="container">
    <div class="stats-grid">
      <div class="stat-card"><h3>Pending Approvals</h3><div class="number">0</div></div>
      <div class="stat-card"><h3>Active Complaints</h3><div class="number">0</div></div>
      <div class="stat-card"><h3>Approved Residents</h3><div class="number">0</div></div>
      <div class="stat-card"><h3>Resolved</h3><div class="number">0</div></div>
    </div>

    <div class="section">
      <h2>Send Announcement to All Residents</h2>
      <div class="form-group">
        <label>Title</label>
        <input type="text" id="announcementTitle" placeholder="e.g., Water Supply Maintenance">
      </div>
      <div class="form-group">
        <label>Message</label>
        <textarea id="announcementMessage" placeholder="Type your message here..."></textarea>
      </div>
      <button class="btn" onclick="sendAnnouncement()">Broadcast to All Residents</button>
    </div>

    <div class="section">
      <h2>Emergency Alerts</h2>
      <div class="empty-state">
        <h3>No active emergencies</h3>
        <p>Emergency alerts will appear here with HIGH priority</p>
      </div>
    </div>

    <div class="section">
      <h2>Registration Requests</h2>
      <div class="empty-state">
        <h3>No pending requests</h3>
        <p>Registration requests will appear here once Twilio is connected</p>
      </div>
    </div>

    <div class="section">
      <h2>Complaints by Category</h2>
      <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 15px; margin-bottom: 20px;">
        <div style="background: #f9f9f9; padding: 15px; border-radius: 4px; text-align: center;">
          <div style="font-size: 24px; font-weight: 700;">0</div>
          <div style="font-size: 12px; color: #666; margin-top: 5px;">Maintenance</div>
        </div>
        <div style="background: #f9f9f9; padding: 15px; border-radius: 4px; text-align: center;">
          <div style="font-size: 24px; font-weight: 700;">0</div>
          <div style="font-size: 12px; color: #666; margin-top: 5px;">Amenity</div>
        </div>
        <div style="background: #f9f9f9; padding: 15px; border-radius: 4px; text-align: center;">
          <div style="font-size: 24px; font-weight: 700;">0</div>
          <div style="font-size: 12px; color: #666; margin-top: 5px;">Other</div>
        </div>
        <div style="background: #ffe6e6; padding: 15px; border-radius: 4px; text-align: center;">
          <div style="font-size: 24px; font-weight: 700; color: #c33;">0</div>
          <div style="font-size: 12px; color: #c33; margin-top: 5px; font-weight: 600;">Emergencies</div>
        </div>
      </div>
    </div>

    <div class="section">
      <h2>Complaints</h2>
      <div class="empty-state">
        <h3>No active complaints</h3>
        <p>Complaints will appear here once residents file them</p>
      </div>
    </div>
  </div>

  <script>
    async function sendAnnouncement() {
      const title = document.getElementById('announcementTitle').value;
      const message = document.getElementById('announcementMessage').value;
      
      if (!title || !message) {
        alert('Please fill in both title and message');
        return;
      }

      try {
        const response = await fetch('/api/broadcast-announcement', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ title, message })
        });

        if (response.ok) {
          alert('Announcement sent to all residents!');
          document.getElementById('announcementTitle').value = '';
          document.getElementById('announcementMessage').value = '';
        }
      } catch (error) {
        alert('Error sending announcement');
        console.error(error);
      }
    }
  </script>
</body>
</html>`;

  res.send(html);
});

// Approve Resident
app.post('/api/approve', async (req, res) => {
  const { flat, phone } = req.body;

  try {
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
        `Your registration for Flat ${flat} has been approved. You can now submit complaints and inquiries.`
      );
    }

    res.status(200).send('OK');
  } catch (error) {
    console.error('Error approving resident:', error);
    res.status(500).send('Error');
  }
});

// Resolve Complaint
app.post('/api/resolve', async (req, res) => {
  const { flat } = req.body;

  try {
    const complaints = await complaintsSheet.getRows();
    const complaint = complaints.find(c => c.Flat === flat && c.Status !== 'Resolved');

    if (complaint) {
      complaint.Status = 'Resolved';
      await complaint.save();

      await sendWhatsAppMessage(
        complaint.Phone,
        `Your complaint for Flat ${flat} has been resolved. Thank you for reporting.`
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