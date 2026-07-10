// index.js — Servidor HTTP que recebe as mensagens do Teams (via Azure Bot Service)
// e as entrega ao bot. Endpoint: POST /api/messages

require('dotenv').config();

const express = require('express');
const {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication
} = require('botbuilder');
const { OnboardingBot } = require('./bot');

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(process.env);
const adapter = new CloudAdapter(botFrameworkAuthentication);

adapter.onTurnError = async (context, error) => {
  console.error('[onTurnError]', error);
  try {
    await context.sendActivity('Ocorreu um erro interno. Tente novamente.');
  } catch { /* ignora */ }
};

const bot = new OnboardingBot();

const app = express();
app.use(express.json());

// Health check (útil para o Azure)
app.get('/', (_req, res) => res.send('Cassotis Onboarding Bot — online'));

app.post('/api/messages', (req, res) => {
  adapter.process(req, res, (context) => bot.run(context));
});

const port = process.env.PORT || 3978;
app.listen(port, () => {
  console.log(`Bot rodando na porta ${port}`);
});
