// bot.js — Lógica do bot no Teams.
// Cada conversa do Teams recebe um session-id próprio na Zoho,
// assim o agente lembra o contexto de cada funcionário separadamente.

const { TeamsActivityHandler, TurnContext } = require('botbuilder');
const { randomUUID } = require('crypto');
const { askAgent } = require('./zohoClient');

// Mapa conversa do Teams -> session-id da Zoho (em memória).
// Se o serviço reiniciar, uma nova sessão é criada — o agente apenas "esquece"
// o histórico daquela conversa, sem causar erro.
const sessions = new Map();

const WELCOME =
  'Olá! 👋 Sou o **Assistente de Onboarding da Cassotis Consulting**.\n\n' +
  'Posso te ajudar com dúvidas sobre processos internos, sistemas, POPs e o seu dia a dia na empresa. ' +
  'É só me perguntar — por exemplo: *"Como criar um POP?"* ou *"Como funciona o booking no MS Bookings?"*';

class OnboardingBot extends TeamsActivityHandler {
  constructor() {
    super();

    // Mensagem recebida de um usuário
    this.onMessage(async (context, next) => {
      const text = TurnContext.removeRecipientMention(context.activity)?.trim()
        || (context.activity.text || '').trim();

      if (!text) {
        await next();
        return;
      }

      const conversationId = context.activity.conversation.id;
      if (!sessions.has(conversationId)) {
        sessions.set(conversationId, randomUUID().replace(/-/g, ''));
      }
      const sessionId = sessions.get(conversationId);

      // Indicador de "digitando..." enquanto a Zoho processa
      await context.sendActivity({ type: 'typing' });

      try {
        const answer = await askAgent(text, sessionId);
        await context.sendActivity(answer);
      } catch (err) {
        console.error('Erro ao consultar o agente Zoho:', err);
        await context.sendActivity(
          'Desculpe, tive um problema para consultar a base de conhecimento agora. ' +
          'Tente novamente em instantes. Se o problema persistir, avise o time de TI.'
        );
      }

      await next();
    });

    // Boas-vindas quando o usuário abre o chat com o bot pela primeira vez
    this.onMembersAdded(async (context, next) => {
      for (const member of context.activity.membersAdded || []) {
        if (member.id !== context.activity.recipient.id) {
          await context.sendActivity(WELCOME);
        }
      }
      await next();
    });
  }
}

module.exports = { OnboardingBot };
