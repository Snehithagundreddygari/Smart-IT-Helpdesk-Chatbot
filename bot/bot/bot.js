import fetch from 'node-fetch';
import { ActivityHandler } from 'botbuilder';
import { matchIntent, faqAnswer } from './recognizers.js';

const TICKET_API = process.env.FUNCTIONS_BASE_URL || 'http://localhost:7071';

export class HelpdeskBot extends ActivityHandler {
  constructor(conversationState){
    super();
    this.conversationState = conversationState;
    this.stateAccessor = this.conversationState.createProperty('dialog');

    this.onMessage(async (context, next) => {
      const text = (context.activity.text || '').trim();
      const intent = matchIntent(text);

      if(intent.name === 'createTicket'){
        await context.sendActivity('Sure — let me raise a ticket. Please provide a short subject for your issue.');
        await this.stateAccessor.set(context, { step: 'subject' });
      }
      else if(intent.name === 'checkTickets'){
        const email = intent.entities.email || 'user@example.com';
        const res = await fetch(`${TICKET_API}/api/GetTickets?email=${encodeURIComponent(email)}`);
        const data = await res.json();
        if(!data || data.length===0){
          await context.sendActivity(`No tickets found for **${email}**.`);
        } else {
          const items = data.slice(0,5).map(t=>`#${t.RowKey} — ${t.Subject} [${t.Status}]`).join('\n');
          await context.sendActivity(`Here are your latest tickets:\n${items}`);
        }
      }
      else if(intent.name === 'faq'){
        const ans = await faqAnswer(text);
        await context.sendActivity(ans);
      }
      else {
        // continue dialog if in progress
        const dlg = await this.stateAccessor.get(context, {});
        if(dlg.step === 'subject'){
          dlg.subject = text; dlg.step = 'description';
          await context.sendActivity('Got it. Please describe the issue in a sentence or two.');
        } else if(dlg.step === 'description'){
          dlg.description = text; dlg.step = 'category';
          await context.sendActivity('Choose a category: Hardware, Software, Access, Network, Other.');
        } else if(dlg.step === 'category'){
          dlg.category = text; dlg.step = 'priority';
          await context.sendActivity('Priority? Low, Medium, or High.');
        } else if(dlg.step === 'priority'){
          dlg.priority = text; dlg.step = undefined;
          // create ticket via Functions
          const payload = {
            subject: dlg.subject,
            description: dlg.description,
            category: dlg.category,
            priority: dlg.priority,
            requesterEmail: context.activity.from?.aadObjectId ? `${context.activity.from.aadObjectId}@contoso.com` : 'user@example.com',
            requesterName: context.activity.from?.name || 'Chat User'
          };
          const res = await fetch(`${TICKET_API}/api/CreateTicket`, {
            method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload)
          });
          const data = await res.json();
          await context.sendActivity(`Ticket created: **#${data.id}** — ${data.subject} [${data.status}]`);
        } else {
          await context.sendActivity("I can help with FAQs, create tickets, or check your tickets. Try: 'raise a ticket', 'my tickets', or ask 'how to reset password?'");
        }
        await this.stateAccessor.set(context, dlg);
      }

      await next();
    });

    this.onMembersAdded(async (context, next) => {
      await context.sendActivity('Hello! I am your IT Helpdesk bot. Say "raise a ticket" to begin, or ask an IT question.');
      await next();
    });
  }
}
