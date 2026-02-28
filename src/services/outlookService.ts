import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';

export class OutlookService {
  private client: Client;

  constructor(accessToken: string) {
    this.client = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
    });
  }

  async listMessages(top: number = 10) {
    return await this.client
      .api('/me/messages')
      .top(top)
      .select('subject,from,receivedDateTime,bodyPreview,isRead')
      .get();
  }

  async getMessage(id: string) {
    return await this.client
      .api(`/me/messages/${id}`)
      .get();
  }

  async sendMessage(subject: string, content: string, to: string) {
    const message = {
      subject,
      body: {
        contentType: 'Text',
        content,
      },
      toRecipients: [
        {
          emailAddress: {
            address: to,
          },
        },
      ],
    };

    return await this.client
      .api('/me/sendMail')
      .post({ message });
  }

  async listEvents(start: string, end: string) {
    return await this.client
      .api('/me/calendarview')
      .query({
        startDateTime: start,
        endDateTime: end,
      })
      .select('subject,start,end,location')
      .get();
  }

  async createEvent(subject: string, start: string, end: string, location?: string) {
    const event = {
      subject,
      start: {
        dateTime: start,
        timeZone: 'UTC',
      },
      end: {
        dateTime: end,
        timeZone: 'UTC',
      },
      location: location ? { displayName: location } : undefined,
    };

    return await this.client
      .api('/me/events')
      .post(event);
  }

  async searchContacts(query: string) {
    return await this.client
      .api('/me/contacts')
      .filter(`contains(displayName, '${query}') or contains(emailAddresses/any(a:a/address), '${query}')`)
      .get();
  }
}
