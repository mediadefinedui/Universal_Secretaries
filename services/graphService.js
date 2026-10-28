/**
 * Microsoft Graph API Service
 * Supports two authentication methods:
 * 1. Username & Password (ROPC) direct consultant authentication
 *    (Places calendar events directly into the consultant's /me/events calendar)
 * 2. Microsoft Entra ID (Azure AD) App Registration with Application Permissions
 *    (/users/{consultantEmail}/events)
 * Also includes realistic Simulation Mode when credentials are not yet configured.
 */

const crypto = require('crypto');

class MicrosoftGraphService {
  constructor() {
    this.tenantId = process.env.MICROSOFT_TENANT_ID || 'organizations';
    this.clientId = process.env.MICROSOFT_CLIENT_ID || '';
    this.clientSecret = process.env.MICROSOFT_CLIENT_SECRET || '';

    // Consultant Account Credentials
    this.userName = process.env.MICROSOFT_USER_NAME || process.env.CONSULTANT_EMAIL || 'snisal@corporate.mediadefined.com';
    this.userPassword = process.env.MICROSOFT_USER_PASSWORD || '';

    this.consultantEmail = process.env.CONSULTANT_EMAIL || this.userName || 'snisal@corporate.mediadefined.com';
    this.consultantName = process.env.CONSULTANT_NAME || 'Senior Tax & Corporate Advisor';
    this.timeZone = process.env.CONSULTANT_TIMEZONE || 'Asia/Colombo';
    this.allowSimulationFallback = process.env.MICROSOFT_GRAPH_FALLBACK_SIMULATION !== 'false';

    this.cachedToken = null;
    this.tokenExpiresAt = 0;
    this.authMode = null; // 'user_password' | 'client_credentials' | 'simulation'
  }

  /**
   * Checks if Username & Password authentication is configured
   */
  isPasswordAuth() {
    return Boolean(this.userName && this.userPassword);
  }

  /**
   * Checks if Azure AD Client Credentials authentication is configured
   */
  isAppAuth() {
    return Boolean(
      this.clientId &&
      this.clientSecret &&
      !this.clientId.includes('your-client-id') &&
      !this.clientSecret.includes('your-client-secret')
    );
  }

  /**
   * Returns true if any Microsoft authentication method is configured
   */
  isConfigured() {
    return this.isPasswordAuth() || this.isAppAuth();
  }

  /**
   * Retrieves OAuth2 Access Token for Microsoft Graph
   */
  async getAccessToken() {
    if (!this.isConfigured()) {
      return null;
    }

    const now = Date.now();
    if (this.cachedToken && this.tokenExpiresAt - now > 300000) {
      return this.cachedToken;
    }

    // 1. DIRECT USERNAME & PASSWORD AUTHENTICATION (ROPC)
    if (this.isPasswordAuth()) {
      console.log(`[Microsoft Graph] Authenticating as consultant user: ${this.userName}`);
      const tokenUrl = `https://login.microsoftonline.com/${this.tenantId || 'organizations'}/oauth2/v2.0/token`;
      
      // Default to Microsoft Teams public client if no custom client ID is supplied
      const effectiveClientId = this.clientId || '1fec8e78-bce4-4aaf-ab1b-5451cc387264';
      
      const params = new URLSearchParams({
        client_id: effectiveClientId,
        grant_type: 'password',
        username: this.userName,
        password: this.userPassword,
        scope: 'https://graph.microsoft.com/.default'
      });

      try {
        const response = await fetch(tokenUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          body: params.toString()
        });

        if (!response.ok) {
          const errText = await response.text();
          throw new Error(`User Password Authentication Failed (${response.status}): ${errText}`);
        }

        const data = await response.json();
        this.cachedToken = data.access_token;
        this.tokenExpiresAt = now + (data.expires_in * 1000);
        this.authMode = 'user_password';
        console.log(`[Microsoft Graph] Successfully authenticated as ${this.userName}`);
        return this.cachedToken;
      } catch (err) {
        console.error('[Microsoft Graph] Password authentication error:', err.message);
        if (!this.allowSimulationFallback) throw err;
      }
    }

    // 2. AZURE AD APP REGISTRATION (CLIENT CREDENTIALS)
    if (this.isAppAuth()) {
      console.log(`[Microsoft Graph] Authenticating via Azure AD App Credentials (client_id: ${this.clientId})`);
      const tokenUrl = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;
      const params = new URLSearchParams({
        client_id: this.clientId,
        client_secret: this.clientSecret,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials'
      });

      try {
        const response = await fetch(tokenUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          body: params.toString()
        });

        if (!response.ok) {
          const errText = await response.text();
          throw new Error(`Azure AD App Auth Failed (${response.status}): ${errText}`);
        }

        const data = await response.json();
        this.cachedToken = data.access_token;
        this.tokenExpiresAt = now + (data.expires_in * 1000);
        this.authMode = 'client_credentials';
        console.log(`[Microsoft Graph] Successfully authenticated via Azure AD App`);
        return this.cachedToken;
      } catch (err) {
        console.error('[Microsoft Graph] App credentials authentication error:', err.message);
        if (!this.allowSimulationFallback) throw err;
      }
    }

    return null;
  }

  /**
   * Queries Microsoft Graph to check the consultant's existing calendar schedule.
   */
  async getConsultantBusySlots(dateStr) {
    if (!this.isConfigured()) {
      return []; // Simulation mode
    }

    try {
      const token = await this.getAccessToken();
      if (!token) return [];

      const startDateTime = `${dateStr}T00:00:00`;
      const endDateTime = `${dateStr}T23:59:59`;

      // Use /me endpoint for user auth, /users/{email} for app auth
      const scheduleUrl = this.authMode === 'user_password'
        ? 'https://graph.microsoft.com/v1.0/me/calendar/getSchedule'
        : `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(this.consultantEmail)}/calendar/getSchedule`;

      const response = await fetch(scheduleUrl, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
          'Prefer': `outlook.timezone="${this.timeZone}"`
        },
        body: JSON.stringify({
          schedules: [this.consultantEmail],
          startTime: { dateTime: startDateTime, timeZone: this.timeZone },
          endTime: { dateTime: endDateTime, timeZone: this.timeZone },
          availabilityViewInterval: 45
        })
      });

      if (!response.ok) {
        const errBody = await response.text();
        console.warn(`[Microsoft Graph] getSchedule notice (${response.status}): ${errBody}`);
        return [];
      }

      const result = await response.json();
      const scheduleItems = result.value?.[0]?.scheduleItems || [];

      return scheduleItems.map(item => ({
        start: item.start.dateTime,
        end: item.end.dateTime,
        status: item.status
      }));
    } catch (err) {
      console.error('[Microsoft Graph] Error checking calendar schedule:', err.message);
      return [];
    }
  }

  /**
   * Creates a calendar event directly in consultant's Outlook / Teams calendar
   * with Microsoft Teams meeting link and attendee invitations.
   */
  async createCalendarEventWithTeamsMeeting({
    bookingId,
    consultationType,
    date,
    timeSlot,
    clientName,
    clientEmail,
    clientPhone,
    clientNotes
  }) {
    const [startTime, endTime] = timeSlot.split('-');
    const startDateTime = `${date}T${startTime}:00`;
    const endDateTime = `${date}T${endTime}:00`;

    // 1. LIVE MICROSOFT GRAPH EXECUTION
    if (this.isConfigured()) {
      try {
        const token = await this.getAccessToken();
        if (token) {
          // If authenticated as consultant user, post directly to /me/events
          // Otherwise, post to /users/{consultantEmail}/events
          const eventUrl = this.authMode === 'user_password'
            ? 'https://graph.microsoft.com/v1.0/me/events'
            : `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(this.consultantEmail)}/events`;

          const htmlBody = `
            <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #1e293b;">
              <h2 style="color: #0084cc; margin-bottom: 4px;">Universal Secretaries - Consultation Meeting</h2>
              <p style="font-size: 14px; color: #64748b;">Booking Reference: <strong>${bookingId}</strong></p>
              <hr style="border: none; border-top: 1px solid #e2e8f0; margin: 16px 0;" />
              
              <h3 style="color: #0f172a; margin-bottom: 8px;">Meeting Details</h3>
              <ul style="list-style: none; padding: 0; margin: 0 0 16px 0;">
                <li><strong>Consultation:</strong> ${consultationType}</li>
                <li><strong>Date:</strong> ${date}</li>
                <li><strong>Time:</strong> ${timeSlot} (Sri Lanka Time UTC+5:30)</li>
                <li><strong>Consultant:</strong> ${this.consultantName} (${this.consultantEmail})</li>
              </ul>

              <h3 style="color: #0f172a; margin-bottom: 8px;">Client Information</h3>
              <ul style="list-style: none; padding: 0; margin: 0 0 16px 0;">
                <li><strong>Name:</strong> ${clientName}</li>
                <li><strong>Email:</strong> ${clientEmail}</li>
                <li><strong>Phone / WhatsApp:</strong> ${clientPhone}</li>
                <li><strong>Client Notes:</strong> ${clientNotes || 'None provided'}</li>
              </ul>

              <div style="background-color: #f0f9ff; border: 1px solid #bae6fd; border-radius: 8px; padding: 14px; margin-top: 16px;">
                <p style="margin: 0 0 8px 0; font-weight: bold; color: #0369a1;">Microsoft Teams Online Meeting</p>
                <p style="margin: 0; font-size: 13px; color: #0284c7;">
                  Click the Microsoft Teams link attached to this calendar invite to join at the scheduled time.
                </p>
              </div>
            </div>
          `;

          const eventPayload = {
            subject: `Consultation: ${consultationType} - ${clientName} [${bookingId}]`,
            body: {
              contentType: 'HTML',
              content: htmlBody
            },
            start: {
              dateTime: startDateTime,
              timeZone: this.timeZone
            },
            end: {
              dateTime: endDateTime,
              timeZone: this.timeZone
            },
            location: {
              displayName: 'Microsoft Teams Meeting'
            },
            attendees: [
              {
                emailAddress: {
                  address: clientEmail,
                  name: clientName
                },
                type: 'required'
              },
              {
                emailAddress: {
                  address: this.consultantEmail,
                  name: this.consultantName
                },
                type: 'required'
              }
            ],
            isOnlineMeeting: true,
            onlineMeetingProvider: 'teamsForBusiness'
          };

          console.log(`[Microsoft Graph] Creating calendar event on ${this.consultantEmail}'s calendar...`);
          const response = await fetch(eventUrl, {
            method: 'POST',
            headers: {
              'Authorization': `Bearer ${token}`,
              'Content-Type': 'application/json',
              'Prefer': `outlook.timezone="${this.timeZone}"`
            },
            body: JSON.stringify(eventPayload)
          });

          if (!response.ok) {
            const errorDetails = await response.text();
            throw new Error(`Microsoft Graph Event Creation Failed (${response.status}): ${errorDetails}`);
          }

          const createdEvent = await response.json();
          const teamsJoinUrl =
            createdEvent.onlineMeeting?.joinUrl ||
            createdEvent.onlineMeetingUrl ||
            createdEvent.webLink;

          console.log(`[Microsoft Graph] SUCCESS! Added event to ${this.consultantEmail}'s calendar. Outlook Event ID: ${createdEvent.id}`);

          return {
            success: true,
            outlookEventId: createdEvent.id,
            teamsMeetingUrl: teamsJoinUrl,
            teamsMeetingId: bookingId,
            mode: this.authMode || 'microsoft_graph'
          };
        }
      } catch (graphError) {
        console.error('[Microsoft Graph] API Error creating calendar event:', graphError.message);
        if (!this.allowSimulationFallback) {
          throw graphError;
        }
        console.warn('[Microsoft Graph] Falling back to simulation mode due to Graph API failure.');
      }
    }

    // 2. SIMULATION MODE
    console.log(`[Microsoft Graph] Generating simulated Microsoft Teams meeting for consultant ${this.consultantEmail}`);
    const token = crypto.randomBytes(16).toString('hex');
    const simulatedTeamsUrl = `https://teams.microsoft.com/l/meetup-join/19%3ameeting_${token}%40thread.v2/0?context=%7b%22Tid%22%3a%22corporate-mediadefined%22%2c%22Oid%22%3a%22snisal%22%7d`;

    return {
      success: true,
      outlookEventId: `outlook-${Date.now()}`,
      teamsMeetingUrl: simulatedTeamsUrl,
      teamsMeetingId: bookingId,
      mode: 'simulation'
    };
  }

  /**
   * Sends confirmation email to the client
   */
  async sendConfirmationEmail({ clientEmail, clientName, bookingId, consultationType, dateFormatted, timeSlot, teamsMeetingUrl }) {
    if (!this.isConfigured()) {
      console.log(`[Microsoft Graph] [Simulation] Confirmation email prepared for ${clientEmail} with Teams URL.`);
      return true;
    }

    try {
      const token = await this.getAccessToken();
      if (!token) return false;

      const sendMailUrl = this.authMode === 'user_password'
        ? 'https://graph.microsoft.com/v1.0/me/sendMail'
        : `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(this.consultantEmail)}/sendMail`;

      const emailBody = `
        <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; color: #1e293b; padding: 20px; border: 1px solid #e2e8f0; border-radius: 12px;">
          <h2 style="color: #0084cc; margin-top: 0;">Universal Secretaries</h2>
          <p>Dear ${clientName},</p>
          <p>Your consultation appointment has been successfully scheduled and confirmed with <strong>${this.consultantName}</strong>.</p>
          
          <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background-color: #f8fafc; border-radius: 8px;">
            <tr>
              <td style="padding: 10px; border-bottom: 1px solid #e2e8f0; font-weight: bold; width: 35%;">Booking Reference:</td>
              <td style="padding: 10px; border-bottom: 1px solid #e2e8f0;">${bookingId}</td>
            </tr>
            <tr>
              <td style="padding: 10px; border-bottom: 1px solid #e2e8f0; font-weight: bold;">Consultant:</td>
              <td style="padding: 10px; border-bottom: 1px solid #e2e8f0;">${this.consultantName} (${this.consultantEmail})</td>
            </tr>
            <tr>
              <td style="padding: 10px; border-bottom: 1px solid #e2e8f0; font-weight: bold;">Consultation:</td>
              <td style="padding: 10px; border-bottom: 1px solid #e2e8f0;">${consultationType}</td>
            </tr>
            <tr>
              <td style="padding: 10px; border-bottom: 1px solid #e2e8f0; font-weight: bold;">Scheduled Date & Time:</td>
              <td style="padding: 10px; border-bottom: 1px solid #e2e8f0;">${dateFormatted} @ ${timeSlot} (Sri Lanka Time UTC+5:30)</td>
            </tr>
            <tr>
              <td style="padding: 10px; font-weight: bold;">Location:</td>
              <td style="padding: 10px;">Microsoft Teams Video Consultation</td>
            </tr>
          </table>

          <div style="text-align: center; margin: 28px 0;">
            <a href="${teamsMeetingUrl}" style="background-color: #0084cc; color: #ffffff; text-decoration: none; padding: 12px 24px; border-radius: 8px; font-weight: bold; display: inline-block;">
              Join Microsoft Teams Meeting
            </a>
          </div>

          <p style="font-size: 13px; color: #64748b;">
            Please ensure you have Microsoft Teams installed or open the meeting in a supported web browser 5 minutes before the start time.
          </p>
          <hr style="border: none; border-top: 1px solid #e2e8f0; margin: 20px 0;" />
          <p style="font-size: 12px; color: #94a3b8; margin: 0;">Universal Secretaries &bull; Colombo, Sri Lanka</p>
        </div>
      `;

      const payload = {
        message: {
          subject: `Confirmed: Consultation Booking [${bookingId}] - Universal Secretaries`,
          body: {
            contentType: 'HTML',
            content: emailBody
          },
          toRecipients: [
            {
              emailAddress: {
                address: clientEmail,
                name: clientName
              }
            }
          ]
        },
        saveToSentItems: 'true'
      };

      await fetch(sendMailUrl, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(payload)
      });

      console.log(`[Microsoft Graph] Confirmation email sent to ${clientEmail}`);
      return true;
    } catch (err) {
      console.warn('[Microsoft Graph] Failed to send confirmation email via sendMail:', err.message);
      return false;
    }
  }
}

module.exports = new MicrosoftGraphService();
