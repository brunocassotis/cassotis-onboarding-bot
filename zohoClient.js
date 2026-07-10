// zohoClient.js — Cliente da Agent API da Zoho Zia Agents.
// Cuida de: (1) renovar o access token via refresh token; (2) enviar a pergunta ao agente.

const ACCOUNTS_URL = process.env.ZOHO_ACCOUNTS_URL || 'https://accounts.zoho.com';
const AGENT_ENDPOINT = process.env.ZOHO_AGENT_ENDPOINT; // URL exata copiada da aba Integrate > Agent API
const ORG_ID = process.env.ZOHO_ORG_ID;
const CLIENT_ID = process.env.ZOHO_CLIENT_ID;
const CLIENT_SECRET = process.env.ZOHO_CLIENT_SECRET;
const REFRESH_TOKEN = process.env.ZOHO_REFRESH_TOKEN;

// Cache do access token (expira em ~1h; renovamos 5 min antes)
let cachedToken = null;
let tokenExpiresAt = 0;

async function getAccessToken() {
  const now = Date.now();
  if (cachedToken && now < tokenExpiresAt - 5 * 60 * 1000) {
    return cachedToken;
  }
  const params = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    refresh_token: REFRESH_TOKEN,
    grant_type: 'refresh_token'
  });
  const res = await fetch(`${ACCOUNTS_URL}/oauth/v2/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params.toString()
  });
  const data = await res.json();
  if (!res.ok || !data.access_token) {
    throw new Error(`Falha ao renovar token Zoho: ${res.status} ${JSON.stringify(data)}`);
  }
  cachedToken = data.access_token;
  tokenExpiresAt = now + (data.expires_in ? data.expires_in * 1000 : 3600 * 1000);
  return cachedToken;
}

/**
 * Envia uma pergunta ao agente Zoho.
 * @param {string} query    - mensagem do usuário
 * @param {string} sessionId - id de sessão (1 por conversa do Teams, mantém o contexto)
 * @returns {string} resposta do agente
 */
async function askAgent(query, sessionId) {
  const token = await getAccessToken();

  const headers = {
    'Content-Type': 'application/json',
    'Authorization': `Zoho-oauthtoken ${token}`,
    'X-ZIAAGENTS-ORG': ORG_ID,
    'X-ZIAAGENTS-AGENT-SESSION-ID': sessionId
  };
  // Compatibilidade com o formato alternativo da API (headers de agente/versão),
  // caso sejam definidos nas variáveis de ambiente.
  if (process.env.ZOHO_AGENT_ID) headers['X-ZIAAGENTS-AGENT-ID'] = process.env.ZOHO_AGENT_ID;
  if (process.env.ZOHO_AGENT_VERSION_ID) headers['X-ZIAAGENTS-AGENT-VERSION-ID'] = process.env.ZOHO_AGENT_VERSION_ID;

  const body = {
    query,
    systemArgs: {},
    reasoning: false,
    attachments: []
  };

  const res = await fetch(AGENT_ENDPOINT, {
    method: 'POST',
    headers,
    body: JSON.stringify(body)
  });

  const raw = await res.text();
  if (!res.ok) {
    throw new Error(`Agent API retornou ${res.status}: ${raw.slice(0, 500)}`);
  }

  let data;
  try {
    data = JSON.parse(raw);
  } catch {
    return raw; // resposta em texto puro
  }

  // Formato documentado: { "data": { "response": "..." } } — com fallbacks defensivos
  return (
    data?.data?.response ??
    data?.data?.answer ??
    data?.response ??
    data?.answer ??
    data?.output ??
    JSON.stringify(data)
  );
}

module.exports = { askAgent };
