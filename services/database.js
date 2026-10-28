const fs = require('fs');
const path = require('path');

const DATA_DIR = path.join(__dirname, '..', 'data');
const DB_FILE = path.join(DATA_DIR, 'bookings.json');

class Database {
  constructor() {
    this.initialized = false;
    this._ensureFile();
  }

  _ensureFile() {
    try {
      if (!fs.existsSync(DATA_DIR)) {
        fs.mkdirSync(DATA_DIR, { recursive: true });
      }
      if (!fs.existsSync(DB_FILE)) {
        fs.writeFileSync(DB_FILE, JSON.stringify({ bookings: [] }, null, 2), 'utf-8');
      }
      this.initialized = true;
    } catch (err) {
      console.error('[Database] Initialization error:', err);
    }
  }

  _read() {
    try {
      this._ensureFile();
      const raw = fs.readFileSync(DB_FILE, 'utf-8');
      return JSON.parse(raw);
    } catch (err) {
      console.error('[Database] Read error, resetting corrupt file:', err);
      return { bookings: [] };
    }
  }

  _write(data) {
    try {
      this._ensureFile();
      const tempFile = `${DB_FILE}.tmp.${Date.now()}`;
      fs.writeFileSync(tempFile, JSON.stringify(data, null, 2), 'utf-8');
      fs.renameSync(tempFile, DB_FILE);
      return true;
    } catch (err) {
      console.error('[Database] Write error:', err);
      // Fallback direct write
      fs.writeFileSync(DB_FILE, JSON.stringify(data, null, 2), 'utf-8');
      return true;
    }
  }

  getAllBookings() {
    const data = this._read();
    return data.bookings || [];
  }

  getBookingById(id) {
    const bookings = this.getAllBookings();
    return bookings.find(b => b.id === id) || null;
  }

  getBookingsByDate(dateStr) {
    const bookings = this.getAllBookings();
    return bookings.filter(b => b.date === dateStr && b.status !== 'cancelled');
  }

  isSlotBooked(dateStr, timeSlot) {
    const bookings = this.getBookingsByDate(dateStr);
    return bookings.some(b => b.timeSlot === timeSlot);
  }

  createBooking(booking) {
    const data = this._read();
    if (!data.bookings) data.bookings = [];

    // Secondary sanity check inside database write
    const alreadyBooked = data.bookings.some(
      b => b.date === booking.date && b.timeSlot === booking.timeSlot && b.status !== 'cancelled'
    );

    if (alreadyBooked) {
      throw new Error(`Double booking conflict: slot ${booking.timeSlot} on ${booking.date} is already booked.`);
    }

    const record = {
      id: booking.id,
      consultationType: booking.consultationType,
      date: booking.date,
      timeSlot: booking.timeSlot,
      dateFormatted: booking.dateFormatted,
      clientName: booking.clientName,
      clientEmail: booking.clientEmail,
      clientPhone: booking.clientPhone,
      clientNotes: booking.clientNotes || '',
      teamsMeetingUrl: booking.teamsMeetingUrl,
      teamsMeetingId: booking.teamsMeetingId,
      outlookEventId: booking.outlookEventId || null,
      status: booking.status || 'confirmed',
      timeZone: booking.timeZone || 'Asia/Colombo',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    data.bookings.push(record);
    this._write(data);
    return record;
  }

  getBookedSlotsMap(dates = []) {
    const bookings = this.getAllBookings().filter(b => b.status !== 'cancelled');
    const result = {};

    dates.forEach(d => {
      result[d] = [];
    });

    bookings.forEach(b => {
      if (!result[b.date]) {
        result[b.date] = [];
      }
      if (!result[b.date].includes(b.timeSlot)) {
        result[b.date].push(b.timeSlot);
      }
    });

    return result;
  }

  deleteBooking(id) {
    const data = this._read();
    const initialLen = data.bookings.length;
    data.bookings = data.bookings.filter(b => b.id !== id);
    if (data.bookings.length !== initialLen) {
      this._write(data);
      return true;
    }
    return false;
  }
}

module.exports = new Database();
