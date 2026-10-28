/**
 * In-Memory Concurrency Lock Manager
 * Prevents race conditions and simultaneous double bookings for the same date & slot
 */

class BookingLockManager {
  constructor() {
    this.activeLocks = new Map(); // key: "YYYY-MM-DD:SLOT" -> timestamp
    this.LOCK_TIMEOUT_MS = 15000; // 15 seconds max lock lifetime to prevent deadlocks
  }

  _getKey(date, timeSlot) {
    return `${date.trim()}#${timeSlot.trim()}`;
  }

  /**
   * Attempts to acquire an exclusive lock for a date & slot.
   * Returns true if lock acquired, false if slot is currently being processed.
   */
  acquireLock(date, timeSlot) {
    const key = this._getKey(date, timeSlot);
    const now = Date.now();

    if (this.activeLocks.has(key)) {
      const lockTime = this.activeLocks.get(key);
      if (now - lockTime < this.LOCK_TIMEOUT_MS) {
        return false; // Still locked by another concurrent request
      }
    }

    this.activeLocks.set(key, now);
    return true;
  }

  /**
   * Releases the lock for a date & slot.
   */
  releaseLock(date, timeSlot) {
    const key = this._getKey(date, timeSlot);
    this.activeLocks.delete(key);
  }

  /**
   * Executes a critical section with an exclusive lock on the slot.
   * Throws 409 Conflict if lock cannot be acquired.
   */
  async withLock(date, timeSlot, taskFn) {
    const acquired = this.acquireLock(date, timeSlot);
    if (!acquired) {
      const err = new Error(`Time slot ${timeSlot} on ${date} is currently being processed by another customer.`);
      err.status = 409;
      err.code = 'SLOT_LOCKED_CONCURRENT';
      throw err;
    }

    try {
      return await taskFn();
    } finally {
      this.releaseLock(date, timeSlot);
    }
  }
}

module.exports = new BookingLockManager();
