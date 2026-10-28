/**
 * Automated Verification Test Suite for Universal Secretaries
 * Tests:
 * 1. Health check & configuration
 * 2. Slot availability check (initial state)
 * 3. Successful booking with Microsoft Teams & Outlook event creation
 * 4. Slot availability post-booking (verifying date & slot are blocked)
 * 5. Double-booking prevention (re-booking the exact same slot returns 409)
 * 6. Simultaneous / Concurrent booking attempts (race-condition protection)
 * 7. Weekend & validation rejection
 */

const http = require('http');
const path = require('path');
const fs = require('fs');

// Set test environment
process.env.NODE_ENV = 'test';
process.env.PORT = '3099';
process.env.MICROSOFT_GRAPH_FALLBACK_SIMULATION = 'true';

const app = require('../server');
const database = require('../services/database');

let server;
const BASE_URL = 'http://localhost:3099';

function request(method, path, data = null) {
  return new Promise((resolve, reject) => {
    const url = new URL(path, BASE_URL);
    const options = {
      hostname: url.hostname,
      port: url.port,
      path: url.pathname + url.search,
      method: method,
      headers: {
        'Content-Type': 'application/json'
      }
    };

    const req = http.request(options, res => {
      let body = '';
      res.on('data', chunk => body += chunk);
      res.on('end', () => {
        try {
          const parsed = body ? JSON.parse(body) : null;
          resolve({ status: res.statusCode, body: parsed });
        } catch (e) {
          resolve({ status: res.statusCode, body });
        }
      });
    });

    req.on('error', reject);
    if (data) {
      req.write(JSON.stringify(data));
    }
    req.end();
  });
}

async function runTests() {
  console.log('\n======================================================');
  console.log(' Starting Booking & Microsoft Teams Integration Tests');
  console.log('======================================================\n');

  let passed = 0;
  let failed = 0;

  function assert(condition, message) {
    if (condition) {
      console.log(`  ✓ PASS: ${message}`);
      passed++;
    } else {
      console.error(`  ✗ FAIL: ${message}`);
      failed++;
    }
  }

  try {
    // 1. Health Check
    console.log('[Test 1] Checking API Health and Microsoft Graph status...');
    const healthRes = await request('GET', '/api/health');
    assert(healthRes.status === 200, 'Health endpoint responds with status 200');
    assert(healthRes.body.status === 'online', 'Server reports online status');
    assert(healthRes.body.consultant.timeZone === 'Asia/Colombo', 'Consultant time zone is Asia/Colombo');

    // Pick a test weekday (e.g. 2026-11-16 is Monday)
    const testDate = '2026-11-16';
    const testSlot = '08:30-09:15';

    // Clean up any test booking for this test date beforehand
    const allBookings = database.getAllBookings();
    allBookings.forEach(b => {
      if (b.date === testDate) database.deleteBooking(b.id);
    });

    // 2. Initial Slot Availability
    console.log(`\n[Test 2] Checking availability for ${testDate}...`);
    const availRes1 = await request('GET', `/api/availability?date=${testDate}`);
    assert(availRes1.status === 200, 'Availability endpoint responds with 200');
    const initialSlot = availRes1.body.slots.find(s => s.timeSlot === testSlot);
    assert(initialSlot && initialSlot.available === true, `Target slot ${testSlot} is initially available`);

    // 3. Successful Booking with Microsoft Teams Link
    console.log(`\n[Test 3] Submitting first booking for ${testDate} at ${testSlot}...`);
    const bookingPayload = {
      consultationType: 'Company registration',
      date: testDate,
      dateFormatted: 'Mon, 16 Nov 2026',
      timeSlot: testSlot,
      clientName: 'Kasun Perera',
      clientPhone: '+94771234567',
      clientEmail: 'kasun.perera@example.com',
      clientNotes: 'PVT LTD registration consultation with 2 directors'
    };

    const bookRes1 = await request('POST', '/api/bookings', bookingPayload);
    assert(bookRes1.status === 201, 'Booking is confirmed with status 201 Created');
    assert(bookRes1.body.booking.id.startsWith('US-LK-'), 'Generated unique booking reference ID');
    assert(bookRes1.body.booking.teamsMeetingUrl.includes('teams.microsoft.com'), 'Generated valid Microsoft Teams meeting URL');
    assert(bookRes1.body.booking.timeSlot === testSlot, 'Confirmed booking matches selected time slot');
    assert(bookRes1.body.booking.status === 'confirmed', 'Booking status is confirmed');
    const createdId = bookRes1.body.booking.id;

    // 4. Verification that Slot is Now Blocked
    console.log(`\n[Test 4] Re-checking availability after booking to verify slot is blocked...`);
    const availRes2 = await request('GET', `/api/availability?date=${testDate}`);
    const postSlot = availRes2.body.slots.find(s => s.timeSlot === testSlot);
    assert(postSlot && postSlot.available === false, `Slot ${testSlot} is now BLOCKED (available: false)`);
    assert(postSlot.reason === 'Booked by customer', 'Blocked reason indicates booked by customer');

    // 5. Double-Booking Prevention: Attempting to Book the Same Slot Again
    console.log(`\n[Test 5] Attempting to book the SAME slot (${testDate} @ ${testSlot}) with another customer...`);
    const duplicatePayload = {
      consultationType: 'Tax consultation',
      date: testDate,
      dateFormatted: 'Mon, 16 Nov 2026',
      timeSlot: testSlot,
      clientName: 'Nimal Fernando',
      clientPhone: '+94779876543',
      clientEmail: 'nimal.fernando@example.com',
      clientNotes: 'Second customer trying to take the same slot'
    };

    const duplicateRes = await request('POST', '/api/bookings', duplicatePayload);
    assert(duplicateRes.status === 409, 'Double booking rejected with 409 Conflict status');
    assert(duplicateRes.body.error.includes('already booked') || duplicateRes.body.code === 'SLOT_ALREADY_BOOKED', 'Error message clearly indicates slot is already booked');

    // 6. Concurrency / Simultaneous Booking Race-Condition Test
    console.log(`\n[Test 6] Testing simultaneous race-condition booking requests on slot 09:15-10:00...`);
    const concurrentSlot = '09:15-10:00';
    const reqA = request('POST', '/api/bookings', {
      consultationType: 'Bookkeeping',
      date: testDate,
      dateFormatted: 'Mon, 16 Nov 2026',
      timeSlot: concurrentSlot,
      clientName: 'Client Alpha',
      clientPhone: '+94771111111',
      clientEmail: 'alpha@example.com'
    });

    const reqB = request('POST', '/api/bookings', {
      consultationType: 'Tax consultation',
      date: testDate,
      dateFormatted: 'Mon, 16 Nov 2026',
      timeSlot: concurrentSlot,
      clientName: 'Client Beta',
      clientPhone: '+94772222222',
      clientEmail: 'beta@example.com'
    });

    const [resA, resB] = await Promise.all([reqA, reqB]);
    const statuses = [resA.status, resB.status];
    const has201 = statuses.includes(201);
    const has409 = statuses.includes(409);

    assert(has201 && has409, `Exactly one concurrent request succeeded (201) and the other was rejected (409) [Statuses: ${statuses.join(', ')}]`);

    // 7. Input Validation & Weekend Rejection
    console.log(`\n[Test 7] Testing weekend rejection (Saturday 2026-11-21)...`);
    const weekendRes = await request('POST', '/api/bookings', {
      consultationType: 'Company registration',
      date: '2026-11-21', // Saturday
      timeSlot: '07:00-07:45',
      clientName: 'Weekend Test',
      clientPhone: '+94770000000',
      clientEmail: 'weekend@example.com'
    });
    assert(weekendRes.status === 400, 'Weekend booking rejected with 400 Bad Request');

    // Cleanup test bookings
    database.deleteBooking(createdId);
    if (resA.body?.booking?.id) database.deleteBooking(resA.body.booking.id);
    if (resB.body?.booking?.id) database.deleteBooking(resB.body.booking.id);

  } catch (err) {
    console.error('Test execution error:', err);
    failed++;
  } finally {
    console.log('\n------------------------------------------------------');
    console.log(` Tests Finished: ${passed} Passed, ${failed} Failed`);
    console.log('------------------------------------------------------\n');
    if (server) server.close();
    process.exit(failed > 0 ? 1 : 0);
  }
}

// Start test server and run suite
server = app.listen(3099, () => {
  runTests();
});

