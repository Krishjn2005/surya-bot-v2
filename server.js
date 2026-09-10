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
let registrationSheet, approvedSheet, complaintsSheet;

async function initializeSheets() {
  try {
    doc = new GoogleSpreadsheet(SHEET_ID);
    await doc.useServiceAccountAuth(GOOGLE_SERVICE_ACCOUNT);
    await doc.loadInfo();

    registrationSheet = doc.sheetsByTitle['Registrations'] || await doc.addSheet({ title: 'Registrations' });
    approvedSheet = doc.sheetsByTitle['Approved Residents'] || await doc.addSheet({ title: 'Approved Residents' });
    complaintsSheet = doc.sheetsByTitle['Complaints'] || await doc.addSheet({ title: 'Complaints' });

    if (registrationSheet.rowCount === 1 && !registrationSheet.headerValues.length) {
      await registrationSheet.setHeaderRow(['Phone', 'Name', 'Flat', 'Status', 'Timestamp']);
    }
    if (approvedSheet.rowCount === 1 && !approvedSheet.headerValues.length) {
      await approvedSheet.setHeaderRow(['Phone', 'Name', 'Flat', 'Approved Date']);
    }
    if (complaintsSheet.rowCount === 1 && !complaintsSheet.headerValues.length) {
      await complaintsSheet.setHeaderRow(['Flat', 'Phone', 'Name', 'Issue Type', 'Description', 'Status', 'Timestamp']);
    }

    console.log('Google Sheets initialized successfully');
  } catch (error) {
    console.error('Error initializing sheets:', error.message);
  }
}

// initializeSheets();

app.post('/whatsapp', async (req, res) => {
  const from = req.body.From.replace('whatsapp:', '');
  const messageBody = req.body.Body.trim();

  try {
    const approvedRows = await approvedSheet.getRows();
    const isApproved = approvedRows.some(row => row.Phone === from);

    if (isApproved) {
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
  if (lowerMsg.includes('clubhouse') || lowerMsg.includes('parking') || lowerMsg.includes('gym')) {
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

app.get('/dashboard', async (req, res) => {
  try {
    const registrations = await registrationSheet.getRows();
    const approved = await approvedSheet.getRows();
    const complaints = await complaintsSheet.getRows();

    const pendingRegs = registrations.filter(r => r.Status === 'Pending');
    const activeComplaints = complaints.filter(c => c.Status !== 'Resolved');
    const resolvedComplaints = complaints.filter(c => c.Status === 'Resolved');

    const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Surya Prakash Residency - Manager Dashboard</title>
  <style>
    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
    }

    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Helvetica Neue', sans-serif;
      background-color: #f5f5f5;
      color: #1a1a1a;
      line-height: 1.6;
    }

    .navbar {
      background-color: #000;
      color: #fff;
      padding: 20px 40px;
      border-bottom: 1px solid #333;
    }

    .navbar h1 {
      font-size: 24px;
      font-weight: 600;
      letter-spacing: 0.5px;
    }

    .navbar p {
      font-size: 13px;
      color: #aaa;
      margin-top: 4px;
    }

    .container {
      max-width: 1200px;
      margin: 0 auto;
      padding: 40px 20px;
    }

    .stats-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
      gap: 20px;
      margin-bottom: 40px;
    }

    .stat-card {
      background: #fff;
      padding: 25px;
      border-radius: 8px;
      border-left: 4px solid #000;
      box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
    }

    .stat-card h3 {
      font-size: 12px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 1px;
      color: #666;
      margin-bottom: 12px;
    }

    .stat-card .number {
      font-size: 36px;
      font-weight: 700;
      color: #000;
    }

    .section {
      background: #fff;
      border-radius: 8px;
      padding: 30px;
      margin-bottom: 30px;
      box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
    }

    .section h2 {
      font-size: 18px;
      font-weight: 600;
      margin-bottom: 25px;
      color: #000;
      letter-spacing: 0.3px;
    }

    table {
      width: 100%;
      border-collapse: collapse;
    }

    thead {
      background-color: #f9f9f9;
      border-bottom: 2px solid #e0e0e0;
    }

    th {
      padding: 14px 15px;
      text-align: left;
      font-weight: 600;
      font-size: 13px;
      color: #333;
      letter-spacing: 0.4px;
      text-transform: uppercase;
    }

    td {
      padding: 14px 15px;
      border-bottom: 1px solid #e9e9e9;
      font-size: 14px;
    }

    tr:hover {
      background-color: #f9f9f9;
    }

    .btn {
      padding: 8px 14px;
      border: none;
      border-radius: 4px;
      font-size: 12px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.2s ease;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      margin-right: 6px;
    }

    .btn-approve {
      background-color: #000;
      color: #fff;
    }

    .btn-approve:hover {
      background-color: #333;
    }

    .btn-resolve {
      background-color: #000;
      color: #fff;
    }

    .btn-resolve:hover {
      background-color: #333;
    }

    .empty-state {
      text-align: center;
      padding: 50px 20px;
      color: #999;
    }

    .empty-state h3 {
      font-size: 16px;
      font-weight: 500;
      color: #666;
      margin-bottom: 8px;
    }

    .empty-state p {
      font-size: 13px;
      color: #999;
    }

    .status-badge {
      display: inline-block;
      padding: 5px 10px;
      border-radius: 3px;
      font-size: 11px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.4px;
      background-color: #f0f0f0;
      color: #666;
    }

    .modal {
      display: none;
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      background: rgba(0, 0, 0, 0.5);
      justify-content: center;
      align-items: center;
      z-index: 1000;
    }

    .modal.active {
      display: flex;
    }

    .modal-content {
      background: #fff;
      padding: 30px;
      border-radius: 8px;
      max-width: 400px;
      box-shadow: 0 10px 40px rgba(0, 0, 0, 0.2);
    }

    .modal-content h3 {
      font-size: 18px;
      font-weight: 600;
      margin-bottom: 12px;
      color: #000;
    }

    .modal-content p {
      font-size: 14px;
      color: #666;
      margin-bottom: 20px;
    }

    .modal-buttons {
      display: flex;
      gap: 10px;
    }

    .modal-buttons button {
      flex: 1;
      padding: 10px;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      font-weight: 600;
      font-size: 13px;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }

    .modal-buttons .confirm {
      background-color: #000;
      color: #fff;
    }

    .modal-buttons .confirm:hover {
      background-color: #333;
    }

    .modal-buttons .cancel {
      background-color: #e0e0e0;
      color: #333;
    }

    .modal-buttons .cancel:hover {
      background-color: #d0d0d0;
    }

    @media (max-width: 768px) {
      .navbar {
        padding: 15px 20px;
      }

      .container {
        padding: 20px 15px;
      }

      .section {
        padding: 20px;
      }

      .stats-grid {
        grid-template-columns: repeat(2, 1fr);
        gap: 15px;
      }

      th, td {
        padding: 10px;
        font-size: 12px;
      }

      .btn {
        padding: 6px 10px;
        font-size: 11px;
      }
    }
  </style>
</head>
<body>
  <div class="navbar">
    <h1>Surya Prakash Residency</h1>
    <p>Manager Dashboard</p>
  </div>

  <div class="container">
    <div class="stats-grid">
      <div class="stat-card">
        <h3>Pending Approvals</h3>
        <div class="number">${pendingRegs.length}</div>
      </div>
      <div class="stat-card">
        <h3>Active Complaints</h3>
        <div class="number">${activeComplaints.length}</div>
      </div>
      <div class="stat-card">
        <h3>Approved Residents</h3>
        <div class="number">${approved.length}</div>
      </div>
      <div class="stat-card">
        <h3>Resolved This Month</h3>
        <div class="number">${resolvedComplaints.length}</div>
      </div>
    </div>

    <div class="section">
      <h2>Registration Requests</h2>
      <table>
        <thead>
          <tr>
            <th>Name</th>
            <th>Flat</th>
            <th>Phone</th>
            <th>Date</th>
            <th>Action</th>
          </tr>
        </thead>
        <tbody>
          ${pendingRegs.length === 0 ? `
            <tr>
              <td colspan="5" class="empty-state">
                <h3>No pending requests</h3>
                <p>Registration requests will appear here</p>
              </td>
            </tr>
          ` : pendingRegs.map(reg => `
            <tr>
              <td>${reg.Name}</td>
              <td>${reg.Flat}</td>
              <td>${reg.Phone}</td>
              <td>${new Date(reg.Timestamp).toLocaleDateString()}</td>
              <td>
                <button class="btn btn-approve" onclick="approveResident('${reg.Flat}', '${reg.Phone}', '${reg.Name}')">Approve</button>
              </td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    </div>

    <div class="section">
      <h2>Complaints</h2>
      <table>
        <thead>
          <tr>
            <th>Flat</th>
            <th>Resident</th>
            <th>Issue</th>
            <th>Type</th>
            <th>Date</th>
            <th>Action</th>
          </tr>
        </thead>
        <tbody>
          ${activeComplaints.length === 0 ? `
            <tr>
              <td colspan="6" class="empty-state">
                <h3>No active complaints</h3>
                <p>Complaints will appear here</p>
              </td>
            </tr>
          ` : activeComplaints.map(comp => `
            <tr>
              <td>${comp.Flat}</td>
              <td>${comp.Name}</td>
              <td>${comp.Description}</td>
              <td>${comp['Issue Type']}</td>
              <td>${new Date(comp.Timestamp).toLocaleDateString()}</td>
              <td>
                <button class="btn btn-resolve" onclick="resolveComplaint('${comp.Flat}', '${comp.Name}')">Resolve</button>
              </td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    </div>
  </div>

  <div class="modal" id="confirmModal">
    <div class="modal-content">
      <h3 id="modalTitle">Confirm Action</h3>
      <p id="modalMessage"></p>
      <div class="modal-buttons">
        <button class="confirm" onclick="confirmAction()">Confirm</button>
        <button class="cancel" onclick="closeModal()">Cancel</button>
      </div>
    </div>
  </div>

  <script>
    let pendingAction = null;

    function approveResident(flat, phone, name) {
      pendingAction = { type: 'approve', flat, phone, name };
      document.getElementById('modalTitle').textContent = 'Approve Registration';
      document.getElementById('modalMessage').textContent = 'Approve ' + name + ' (Flat ' + flat + ')?';
      document.getElementById('confirmModal').classList.add('active');
    }

    function resolveComplaint(flat, name) {
      pendingAction = { type: 'resolve', flat, name };
      document.getElementById('modalTitle').textContent = 'Resolve Complaint';
      document.getElementById('modalMessage').textContent = 'Mark complaint for Flat ' + flat + ' as resolved?';
      document.getElementById('confirmModal').classList.add('active');
    }

    function confirmAction() {
      if (!pendingAction) return;

      if (pendingAction.type === 'approve') {
        fetch('/api/approve', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ flat: pendingAction.flat, phone: pendingAction.phone })
        }).then(() => {
          closeModal();
          setTimeout(() => location.reload(), 800);
        });
      } else if (pendingAction.type === 'resolve') {
        fetch('/api/resolve', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ flat: pendingAction.flat })
        }).then(() => {
          closeModal();
          setTimeout(() => location.reload(), 800);
        });
      }
    }

    function closeModal() {
      document.getElementById('confirmModal').classList.remove('active');
      pendingAction = null;
    }
  </script>
</body>
</html>`;

    res.send(html);
  } catch (error) {
    console.error('Dashboard error:', error);
    res.status(500).send('Error loading dashboard');
  }
});

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
        `Your registration for Flat ${flat} has been approved. You can now submit complaints and inquiries through this service.`
      );
    }

    res.status(200).send('OK');
  } catch (error) {
    console.error('Error approving resident:', error);
    res.status(500).send('Error');
  }
});

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