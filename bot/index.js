import 'dotenv/config';
import express from 'express';
import { BotFrameworkAdapter, MemoryStorage, ConversationState } from 'botbuilder';
import { HelpdeskBot } from './bot.js';

const app = express();
const adapter = new BotFrameworkAdapter({
  appId: process.env.MICROSOFT_APP_ID || '',
  appPassword: process.env.MICROSOFT_APP_PASSWORD || ''
});

adapter.onTurnError = async (context, error) => {
  console.error('Bot Error:', error);
  await context.sendActivity('Sorry, something went wrong.');
};

const memory = new MemoryStorage();
const conversationState = new ConversationState(memory);
const bot = new HelpdeskBot(conversationState);

app.post('/api/messages', (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
});

const port = process.env.PORT || 3978;
app.listen(port, () => console.log(`Bot listening on :${port}`));
