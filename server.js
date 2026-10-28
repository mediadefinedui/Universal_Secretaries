require('dotenv').config();
const express = require('express');
const cors = require('cors');
const path = require('path');

const database = require('./services/database');
const bookingLock = require('./services/bookingLock');
const graphService = require('./services/graphService');

const app = express();
const PORT = process.env.PORT || 3000;

// Standard consultant time slots (Sri Lanka Time UTC+5:30)
const STANDARD_SLOTS = [
  '07:00-07:45',
  '07:45-08:30',
  '08:30-09:15',
  '09:15-10:00',
  '10:00-10:45',
  '10:45-11:30'
];

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Serve static frontend files from the current directory
app.use(express.static(__dirname));

/**
 * Health & Configuration Status Check
 */
app.get('/api/health', (req, res) => {
  const isGraphConfigured = graphService.isConfigured();
  const bookings = database.getAllBookings();

  res.json({
    status: 'online',
    timestamp: new Date().toISOString(),
    service: 'Universal Secretaries Consultation Booking API',
    consultant: {
      name: process.env.CONSULTANT_NAME || 'Senior Tax & Corporate Advisor',
      email: process.env.CONSULTANT_EMAIL || 'snisal@corporate.mediadefined.com',
      timeZone: process.env.CONSULTANT_TIMEZONE || 'Asia/Colombo'
    },
    microsoftGraph: {
      connected: isGraphConfigured,
      mode: isGraphConfigured ? 'live_microsoft_365' : 'simulation_fallback'
    },
    totalBookings: bookings.length
  });
});

/**
 * Get availability for a specific date (combining DB bookings + Microsoft Graph Calendar)
 * GET /api/availability?date=YYYY-MM-DD
 */
app.get('/api/availability', async (req, res) => {
  try {
    const { date } = req.query;

    if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      return res.status(400).json({ error: 'Valid date parameter (YYYY-MM-DD) is required.' });
    }

    // 1. Get confirmed bookings from database
    const dbBookings = database.getBookingsByDate(date);
    const bookedInDb = new Set(dbBookings.map(b => b.timeSlot));

    // 2. Query Microsoft Graph calendar for consultant busy times
    let graphBusySlots = [];
    try {
      graphBusySlots = await graphService.getConsultantBusySlots(date);
    } catch (gErr) {
      console.warn('[Server] Could not fetch Graph calendar schedule:', gErr.message);
    }

    // 3. Map slot availability
    const slots = STANDARD_SLOTS.map(slot => {
      let isAvailable = true;
      let reason = null;

      if (bookedInDb.has(slot)) {
        isAvailable = false;
        reason = 'Booked by customer';
      } else {
        // Check if Graph schedule marks this interval as busy
        const [startStr, endStr] = slot.split('-');
        const slotStart = new Date(`${date}T${startStr}:00+05:30`).getTime();
        const slotEnd = new Date(`${date}T${endStr}:00+05:30`).getTime();

        const overlaps = graphBusySlots.some(b => {
          const bStart = new Date(b.start).getTime();
          const bEnd = new Date(b.end).getTime();
          return slotStart < bEnd && slotEnd > bStart;
        });

        if (overlaps) {
          isAvailable = false;
          reason = 'Consultant busy in Outlook calendar';
        }
      }

      return {
        timeSlot: slot,
        available: isAvailable,
        reason: reason
      };
    });

    const availableCount = slots.filter(s => s.available).length;
    const fullyBooked = availableCount === 0;

    res.json({
      date,
      timeZone: process.env.CONSULTANT_TIMEZONE || 'Asia/Colombo',
      totalSlots: STANDARD_SLOTS.length,
      availableCount,
      fullyBooked,
      slots
    });
  } catch (err) {
    console.error('[Server] Availability check error:', err);
    res.status(500).json({ error: 'Failed to check slot availability.' });
  }
});

/**
 * Batch availability check for upcoming dates (for calendar day cards)
 * GET /api/availability/batch?dates=YYYY-MM-DD,YYYY-MM-DD,...
 */
app.get('/api/availability/batch', (req, res) => {
  try {
    const { dates } = req.query;
    if (!dates) {
      return res.status(400).json({ error: 'dates query parameter is required (comma separated).' });
    }

    const dateList = dates.split(',').map(d => d.trim()).filter(d => /^\d{4}-\d{2}-\d{2}$/.test(d));
    const bookedMap = database.getBookedSlotsMap(dateList);

    const summary = {};
    dateList.forEach(d => {
      const booked = bookedMap[d] || [];
      const availableCount = STANDARD_SLOTS.length - booked.length;
      summary[d] = {
        fullyBooked: availableCount <= 0,
        availableCount: Math.max(0, availableCount),
        bookedSlots: booked
      };
    });

    res.json({ availability: summary });
  } catch (err) {
    console.error('[Server] Batch availability error:', err);
    res.status(500).json({ error: 'Failed to retrieve batch availability.' });
  }
});

/**
 * Create a new consultation booking
 * Prevents double booking via Concurrency Mutex Lock,
 * creates Microsoft Teams Meeting & Outlook Calendar event via Graph API,
 * persists booking to DB, and sends email confirmation.
 * POST /api/bookings
 */
app.post('/api/bookings', async (req, res) => {
  const {
    consultationType,
    date,
    dateFormatted,
    timeSlot,
    clientName,
    clientPhone,
    clientEmail,
    clientNotes
  } = req.body;

  // Validation
  if (!consultationType || !date || !timeSlot || !clientName || !clientPhone || !clientEmail) {
    return res.status(400).json({
      error: 'Missing required booking details (consultationType, date, timeSlot, clientName, clientPhone, clientEmail).'
    });
  }

  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    return res.status(400).json({ error: 'Invalid date format. Expected YYYY-MM-DD.' });
  }

  if (!STANDARD_SLOTS.includes(timeSlot)) {
    return res.status(400).json({ error: `Invalid time slot: ${timeSlot}.` });
  }

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(clientEmail)) {
    return res.status(400).json({ error: 'Invalid email address provided.' });
  }

  // Prevent bookings on weekends
  const dayOfWeek = new Date(`${date}T12:00:00Z`).getUTCDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    return res.status(400).json({ error: 'Bookings are only available on weekdays (Monday - Friday).' });
  }

  try {
    // Execute inside atomic concurrency lock for date+slot
    const bookingResult = await bookingLock.withLock(date, timeSlot, async () => {
      // 1. RE-VALIDATE AVAILABILITY ON BACKEND (DOUBLE-BOOKING GUARD)
      if (database.isSlotBooked(date, timeSlot)) {
        const conflictErr = new Error(`The slot ${timeSlot} on ${date} is already booked.`);
        conflictErr.status = 409;
        conflictErr.code = 'SLOT_ALREADY_BOOKED';
        throw conflictErr;
      }

      // Check live Graph availability
      const busySlots = await graphService.getConsultantBusySlots(date);
      const [startStr, endStr] = timeSlot.split('-');
      const slotStart = new Date(`${date}T${startStr}:00+05:30`).getTime();
      const slotEnd = new Date(`${date}T${endStr}:00+05:30`).getTime();

      const overlaps = busySlots.some(b => {
        const bStart = new Date(b.start).getTime();
        const bEnd = new Date(b.end).getTime();
        return slotStart < bEnd && slotEnd > bStart;
      });

      if (overlaps) {
        const conflictErr = new Error(`The consultant is busy during ${timeSlot} on ${date} in Microsoft Outlook.`);
        conflictErr.status = 409;
        conflictErr.code = 'CONSULTANT_CALENDAR_BUSY';
        throw conflictErr;
      }

      // 2. GENERATE BOOKING REFERENCE
      const randNum = Math.floor(100000 + Math.random() * 900000);
      const bookingId = `US-LK-${randNum}`;

      // 3. CREATE MICROSOFT OUTLOOK EVENT & TEAMS MEETING LINK VIA GRAPH API
      // If this throws an error, the catch block triggers, the lock is freed,
      // and NO booking is saved to the database.
      const graphResult = await graphService.createCalendarEventWithTeamsMeeting({
        bookingId,
        consultationType,
        date,
        timeSlot,
        clientName: clientName.trim(),
        clientEmail: clientEmail.trim(),
        clientPhone: clientPhone.trim(),
        clientNotes: clientNotes ? clientNotes.trim() : ''
      });

      // 4. PERSIST TO APPLICATION DATABASE
      const savedBooking = database.createBooking({
        id: bookingId,
        consultationType,
        date,
        dateFormatted: dateFormatted || date,
        timeSlot,
        clientName: clientName.trim(),
        clientEmail: clientEmail.trim(),
        clientPhone: clientPhone.trim(),
        clientNotes: clientNotes ? clientNotes.trim() : '',
        teamsMeetingUrl: graphResult.teamsMeetingUrl,
        teamsMeetingId: graphResult.teamsMeetingId || bookingId,
        outlookEventId: graphResult.outlookEventId || null,
        status: 'confirmed',
        timeZone: process.env.CONSULTANT_TIMEZONE || 'Asia/Colombo'
      });

      // 5. SEND CONFIRMATION EMAIL VIA MICROSOFT GRAPH (Non-blocking)
      if (process.env.SEND_CONFIRMATION_EMAIL !== 'false') {
        graphService.sendConfirmationEmail({
          clientEmail: clientEmail.trim(),
          clientName: clientName.trim(),
          bookingId,
          consultationType,
          dateFormatted: dateFormatted || date,
          timeSlot,
          teamsMeetingUrl: graphResult.teamsMeetingUrl
        }).catch(emailErr => {
          console.warn('[Server] Background email notification notice:', emailErr.message);
        });
      }

      return savedBooking;
    });

    return res.status(201).json({
      success: true,
      message: 'Consultation successfully scheduled with Microsoft Teams meeting.',
      booking: bookingResult
    });

  } catch (err) {
    console.error('[Server] Booking creation failed:', err.message);

    if (err.status === 409 || err.code === 'SLOT_ALREADY_BOOKED' || err.code === 'SLOT_LOCKED_CONCURRENT') {
      return res.status(409).json({
        error: err.message || 'This date and time slot has just been booked. Please select another available slot.',
        code: err.code || 'SLOT_CONFLICT'
      });
    }

    return res.status(500).json({
      error: 'Unable to schedule consultation at this time. Please try again or contact consultant directly.',
      details: err.message
    });
  }
});

/**
 * Get all confirmed bookings (for consultant management)
 */
app.get('/api/bookings', (req, res) => {
  const bookings = database.getAllBookings();
  res.json({ count: bookings.length, bookings });
});

/**
 * Get single booking by ID
 */
app.get('/api/bookings/:id', (req, res) => {
  const booking = database.getBookingById(req.params.id);
  if (!booking) {
    return res.status(404).json({ error: 'Booking not found.' });
  }
  res.json({ booking });
});

// Fallback to index.html for root path
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// Start listening if run directly
if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`======================================================================`);
    console.log(` Universal Secretaries Consultation & Microsoft Teams Server`);
    console.log(` Server running on http://localhost:${PORT}`);
    console.log(` Mode: ${graphService.isConfigured() ? 'Microsoft 365 Connected' : 'High-Fidelity Simulation Mode'}`);
    console.log(` Consultant: ${process.env.CONSULTANT_NAME || 'Advisor'} (${process.env.CONSULTANT_EMAIL || 'snisal@corporate.mediadefined.com'})`);
    console.log(` Timezone: ${process.env.CONSULTANT_TIMEZONE || 'Asia/Colombo'}`);
    console.log(`======================================================================`);
  });
}

module.exports = app;
