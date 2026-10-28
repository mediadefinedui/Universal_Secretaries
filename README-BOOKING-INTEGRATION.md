# Universal Secretaries - Consultation Booking & Microsoft 365 / Teams Integration

This document explains the architecture, features, double-booking prevention mechanisms, and Microsoft 365 (Microsoft Teams & Outlook Calendar) integration added to the Universal Secretaries application.

---

## 🌟 Key Features Implemented

### 1. Prevent Double Bookings
- **Dynamic Slot Availability:** When a user selects a date on `index.html`, the application queries the backend API (`/api/availability?date=YYYY-MM-DD`), checking both the local database and the consultant's Microsoft Outlook calendar schedule.
- **Visual Slot Blocking:** Booked slots are automatically disabled with a `.slot-booked` state, strikethrough styling, and a red `"BOOKED"` indicator badge, preventing customer selection.
- **Fully-Booked Day Indicators:** Days where all slots are taken display a `"Full"` tag on the date picker carousel and a warning notice.
- **Backend Concurrency Mutex Lock:** Uses an in-memory transactional lock (`services/bookingLock.js`) around each `date + timeSlot` combination. If two customers attempt to submit the form at the same millisecond, only the first request succeeds; the concurrent request receives a `409 Conflict: SLOT_ALREADY_BOOKED` response.
- **Atomic Rollback on API Failure:** The booking is committed to the database **only** after the calendar event and Teams meeting are successfully confirmed. If Microsoft Graph API fails, the slot is released without creating phantom bookings.

### 2. Microsoft Teams & Outlook Calendar Integration
- **Outlook Calendar Event Creation:** Uses the **Microsoft Graph API** (`POST /users/{consultantEmail}/events`) to generate a confirmed event on the consultant's calendar.
- **Automated Microsoft Teams Meeting:** The calendar event is configured with `isOnlineMeeting: true` and `onlineMeetingProvider: "teamsForBusiness"`, producing a unique, authenticated Microsoft Teams join link.
- **Attendee Invitations:** Both the client and the consultant are invited as required attendees. Microsoft Exchange automatically delivers meeting invitations with Teams join links directly to their Outlook/Exchange inbox.
- **Confirmation Email:** Delivers a branded booking confirmation email with reference ID, consultation details, and Microsoft Teams join button.
- **Confirmation Modal:** Displays the booking reference ID (`US-LK-XXXXXX`), consultation type, scheduled date/time in Sri Lanka time (UTC+5:30), direct Teams join link, WhatsApp notification link, and downloadable `.ics` calendar file.

### 3. Microsoft 365 / Entra ID Security
- Uses secure **OAuth2 Client Credentials Flow** with Microsoft Entra ID (Azure AD).
- Secrets and tenant configurations are stored exclusively on the server in `.env`.
- No client secrets or access tokens are ever exposed to the frontend.
- High-fidelity **Simulation Mode** allows immediate local testing even before configuring live Azure tenant credentials.

---

## 📁 Project Architecture

```
UniversalSecretaries/
├── index.html                           # Main homepage with consultation booking card & modal
├── package.json                         # Node.js dependencies & scripts
├── server.js                            # Express backend API & static file server
├── .env.example                         # Example environment variables
├── .env                                 # Active configuration file
├── data/
│   └── bookings.json                    # Transactional persistent booking store
├── services/
│   ├── database.js                      # ACID file-based database for bookings
│   ├── bookingLock.js                   # In-memory mutex lock for race-condition prevention
│   └── graphService.js                  # Microsoft Graph API client (Entra ID, Calendar, Teams, Mail)
├── assets/
│   ├── js/
│   │   └── booking-consultation.js      # Frontend booking controller & dynamic availability
│   └── css/
│       └── booking-consultation.css     # Styling for slots, blocked badges, and modal
└── test/
    └── booking-test.js                  # Automated verification test suite
```

---

## 🚀 How to Run the Application

### 1. Install Dependencies
```bash
npm install
```
*(Dependencies: `express`, `cors`, `dotenv`)*

### 2. Configure Environment (`.env`)
Inspect or update `.env`:
```ini
PORT=3000
NODE_ENV=development

# Consultant Details (Sri Lanka Time UTC+5:30)
CONSULTANT_NAME="Senior Tax & Corporate Advisor"
CONSULTANT_EMAIL="tax@universalsecretaries.com"
CONSULTANT_PHONE="94774087869"
CONSULTANT_TIMEZONE="Asia/Colombo"

# Microsoft 365 / Entra ID Credentials
MICROSOFT_TENANT_ID=""
MICROSOFT_CLIENT_ID=""
MICROSOFT_CLIENT_SECRET=""

# Fallback Simulation (set to true for offline / dev testing)
MICROSOFT_GRAPH_FALLBACK_SIMULATION=true
SEND_CONFIRMATION_EMAIL=true
```

### 3. Start the Server
```bash
npm start
```
The application will be live at: **`http://localhost:3000`**

### 4. Run Automated Tests
```bash
npm test
```
The test suite validates:
1. API Health and Consultant Timezone (`Asia/Colombo`).
2. Slot availability check (all standard slots available).
3. Booking submission with Microsoft Teams link creation.
4. Slot blocking post-booking (verifying the slot is now blocked).
5. Double-booking prevention (`409 Conflict: Already booked`).
6. Concurrent race-condition handling (simultaneous requests to the same slot).
7. Input validation and weekend booking rejection (`400 Bad Request`).

---

## 🔑 Microsoft 365 / Entra ID (Azure AD) Setup Guide

To connect live with your organization's Microsoft 365 tenant:

1. Sign in to the **[Microsoft Entra Admin Center](https://entra.microsoft.com)**.
2. Navigate to **Identity** > **Applications** > **App registrations** > **New registration**.
3. Name the app (e.g., `Universal Secretaries Booking Integration`) and click **Register**.
4. Note down:
   - **Application (client) ID** -> `MICROSOFT_CLIENT_ID`
   - **Directory (tenant) ID** -> `MICROSOFT_TENANT_ID`
5. Navigate to **API permissions** > **Add a permission** > **Microsoft Graph** > **Application permissions**:
   - `Calendars.ReadWrite` (To inspect calendar schedules and create Outlook events)
   - `OnlineMeetings.ReadWrite.All` (To generate Microsoft Teams meetings)
   - `Mail.Send` (To send confirmation emails to clients)
6. Click **Grant admin consent for [Your Organization]**.
7. Navigate to **Certificates & secrets** > **New client secret**.
8. Copy the generated **Value** -> `MICROSOFT_CLIENT_SECRET`.
9. Paste these credentials into your `.env` file and restart the server.
