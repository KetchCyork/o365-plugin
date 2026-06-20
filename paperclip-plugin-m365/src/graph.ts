/**
 * Graph client (read-only)
 * ------------------------
 * Thin wrapper over Microsoft Graph v1.0. Returns small, structured shapes —
 * subjects, senders, times, names — not full message bodies, unless snippets are
 * explicitly enabled. Keeps work content minimal so only summaries leave the node.
 */
import type { M365Auth } from "./auth.js";
import type { M365Config } from "./config.js";

const GRAPH = "https://graph.microsoft.com/v1.0";

export interface MailItem { id: string; subject: string; from: string; received: string; isRead: boolean; preview?: string; }
export interface EventItem { subject: string; start: string; end: string; organizer: string; location?: string; }
export interface DriveItem { name: string; isFolder: boolean; modified: string; webUrl: string; size?: number; }
export interface ChatItem { topic: string; lastActivity: string; }

export class GraphClient {
  constructor(private auth: M365Auth, private cfg: M365Config) {}

  private async get(path: string): Promise<any> {
    const token = await this.auth.requireToken();
    const res = await fetch(`${GRAPH}${path}`, {
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    });
    if (!res.ok) throw new Error(`Graph ${res.status} on ${path}: ${await res.text()}`);
    return res.json();
  }

  private async send(method: "POST" | "PATCH", path: string, body?: unknown): Promise<any> {
    const token = await this.auth.requireToken();
    const res = await fetch(`${GRAPH}${path}`, {
      method,
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: body === undefined ? undefined : JSON.stringify(body),
    });
    if (!res.ok) throw new Error(`Graph ${method} ${res.status} on ${path}: ${await res.text()}`);
    return res.status === 202 || res.status === 204 ? {} : res.json();
  }

  async me(): Promise<{ displayName: string; mail: string }> {
    const d = await this.get(`/me?$select=displayName,mail,userPrincipalName`);
    return { displayName: d.displayName, mail: d.mail ?? d.userPrincipalName };
  }

  async recentMail(top = 10, unreadOnly = false): Promise<MailItem[]> {
    const select = "id,subject,from,receivedDateTime,isRead,bodyPreview";
    const filter = unreadOnly ? "&$filter=isRead eq false" : "";
    const d = await this.get(
      `/me/mailFolders/inbox/messages?$top=${top}&$orderby=receivedDateTime desc&$select=${select}${filter}`
    );
    return (d.value ?? []).map((m: any): MailItem => ({
      id: m.id,
      subject: m.subject ?? "(no subject)",
      from: m.from?.emailAddress?.name ?? m.from?.emailAddress?.address ?? "(unknown)",
      received: m.receivedDateTime,
      isRead: !!m.isRead,
      preview: this.cfg.allowSnippets ? String(m.bodyPreview ?? "").slice(0, 200) : undefined,
    }));
  }

  async searchMail(query: string, top = 10): Promise<MailItem[]> {
    const select = "id,subject,from,receivedDateTime,isRead,bodyPreview";
    const d = await this.get(
      `/me/messages?$search=${encodeURIComponent(`"${query}"`)}&$top=${top}&$select=${select}`
    );
    return (d.value ?? []).map((m: any): MailItem => ({
      id: m.id,
      subject: m.subject ?? "(no subject)",
      from: m.from?.emailAddress?.name ?? "(unknown)",
      received: m.receivedDateTime,
      isRead: !!m.isRead,
      preview: this.cfg.allowSnippets ? String(m.bodyPreview ?? "").slice(0, 200) : undefined,
    }));
  }

  async events(fromIso: string, toIso: string): Promise<EventItem[]> {
    const select = "subject,start,end,organizer,location";
    const d = await this.get(
      `/me/calendarView?startDateTime=${fromIso}&endDateTime=${toIso}` +
      `&$orderby=start/dateTime&$select=${select}&$top=25`
    );
    return (d.value ?? []).map((e: any): EventItem => ({
      subject: e.subject ?? "(no subject)",
      start: e.start?.dateTime ?? "",
      end: e.end?.dateTime ?? "",
      organizer: e.organizer?.emailAddress?.name ?? "(unknown)",
      location: e.location?.displayName || undefined,
    }));
  }

  async listDrive(path?: string): Promise<DriveItem[]> {
    const endpoint = path
      ? `/me/drive/root:/${encodeURIComponent(path)}:/children`
      : `/me/drive/root/children`;
    const d = await this.get(`${endpoint}?$select=name,size,lastModifiedDateTime,folder,file,webUrl&$top=50`);
    return (d.value ?? []).map((i: any): DriveItem => ({
      name: i.name,
      isFolder: !!i.folder,
      modified: i.lastModifiedDateTime,
      webUrl: i.webUrl,
      size: i.size,
    }));
  }

  async searchDrive(query: string): Promise<DriveItem[]> {
    const d = await this.get(
      `/me/drive/root/search(q='${encodeURIComponent(query)}')?$select=name,lastModifiedDateTime,folder,file,webUrl&$top=25`
    );
    return (d.value ?? []).map((i: any): DriveItem => ({
      name: i.name,
      isFolder: !!i.folder,
      modified: i.lastModifiedDateTime,
      webUrl: i.webUrl,
      size: i.size,
    }));
  }

  /**
   * Recent Teams chats. NOTE: deeper Teams message access (channel messages)
   * often needs admin-consented scopes; this sticks to the user's own chats.
   */
  async recentChats(top = 10): Promise<ChatItem[]> {
    const d = await this.get(`/me/chats?$top=${top}&$select=topic,lastUpdatedDateTime`);
    return (d.value ?? []).map((c: any): ChatItem => ({
      topic: c.topic ?? "(direct chat)",
      lastActivity: c.lastUpdatedDateTime ?? "",
    }));
  }

  // --- Write operations (all create DRAFTS / HOLDS; nothing is sent) ---------

  /** Create a draft email in the Drafts folder. Reversible; not sent. */
  async createDraft(input: { to?: string[]; subject: string; body: string }): Promise<DraftRef> {
    const msg = {
      subject: input.subject,
      body: { contentType: "Text", content: input.body },
      toRecipients: (input.to ?? []).map((a) => ({ emailAddress: { address: a } })),
    };
    const d = await this.send("POST", `/me/messages`, msg);
    return { id: d.id, webLink: d.webLink };
  }

  /** Create a draft reply (or reply-all) to an existing message, with your text. */
  async createReplyDraft(messageId: string, replyText: string, replyAll = false): Promise<DraftRef> {
    const action = replyAll ? "createReplyAll" : "createReply";
    const d = await this.send("POST", `/me/messages/${messageId}/${action}`, { comment: replyText });
    return { id: d.id, webLink: d.webLink };
  }

  /**
   * Create a TENTATIVE calendar hold on your own calendar — no attendees, so no
   * invitations are sent. Proposed attendees + agenda go in the body for you to
   * review, add, and send from Outlook.
   */
  async createCalendarHold(input: {
    subject: string; start: string; end: string;
    agenda?: string; location?: string; proposedAttendees?: string[];
  }): Promise<DraftRef> {
    const bodyText =
      (input.agenda ? `${input.agenda}\n\n` : "") +
      (input.proposedAttendees?.length ? `Proposed attendees: ${input.proposedAttendees.join(", ")}` : "");
    const ev = {
      subject: input.subject,
      start: { dateTime: input.start, timeZone: this.cfg.timeZone },
      end: { dateTime: input.end, timeZone: this.cfg.timeZone },
      body: { contentType: "Text", content: bodyText },
      location: input.location ? { displayName: input.location } : undefined,
      showAs: "tentative",
      isReminderOn: true,
    };
    const d = await this.send("POST", `/me/events`, ev);
    return { id: d.id, webLink: d.webLink };
  }

  /**
   * Send an existing draft. HUMAN-ONLY path: requires the Mail.Send scope and is
   * never invoked by the autonomous executor. The simplest approval is to send
   * from Outlook after reviewing; this exists for an explicit, confirmed command.
   */
  async sendDraft(messageId: string): Promise<void> {
    await this.send("POST", `/me/messages/${messageId}/send`);
  }
}

export interface DraftRef { id: string; webLink: string; }
