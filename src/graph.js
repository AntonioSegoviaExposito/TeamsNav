import { createHash } from 'crypto';
import { mkdir, readFile, rename, writeFile } from 'fs/promises';
import { tmpdir } from 'os';
import { dirname, join } from 'path';

const TENANT = process.env.TENANT || 'organizations';
const CLIENT_ID = process.env.CLIENT_ID || '14d82eec-204b-4c2f-b7e8-296a70dab67e'; // Microsoft Graph Command Line Tools
const GRAPH = 'https://graph.microsoft.com/v1.0';
const LOGIN = `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0`;

// One cache folder per tenant + client: token, chat metadata and message history.
export const CACHE = join(tmpdir(), 'teams-nav', createHash('sha256').update(`${TENANT}|${CLIENT_ID}`).digest('hex').slice(0, 12));

export const CHAT_SCOPES = ['Chat.ReadWrite', 'User.Read'];
export const CHANNEL_SCOPES = [...CHAT_SCOPES, 'ChannelMessage.Read.All', 'Team.ReadBasic.All', 'Channel.ReadBasic.All'];

export const sleep = (ms) => new Promise(r => setTimeout(r, ms));

export async function load(name, fallback) {
  try {
    return JSON.parse(await readFile(join(CACHE, name), 'utf-8'));
  } catch {
    return fallback;
  }
}

// Atomic write: several MCP clients may share the cache folder.
export async function save(name, data) {
  const file = join(CACHE, name);
  await mkdir(dirname(file), { recursive: true, mode: 0o700 });
  const tmp = `${file}.${process.pid}.tmp`;
  await writeFile(tmp, JSON.stringify(data), { mode: 0o600 });
  await rename(tmp, file);
}

// ---- Auth: device-code sign-in once, then silent refresh. One access token per scope set.

const scopeParam = (scopes) => [...scopes.map(s => `https://graph.microsoft.com/${s}`), 'offline_access'].join(' ');
let auth; // { refresh_token, tokens: { [scopes]: { access_token, expires_at } } }
let login; // pending device-code sign-in: { message, expiresAt }
let loginFailure;
const inflight = new Map();

async function oauth(endpoint, params) {
  const res = await fetch(`${LOGIN}/${endpoint}`, { method: 'POST', body: new URLSearchParams({ client_id: CLIENT_ID, ...params }) });
  return res.json();
}

async function storeToken(res, scopes) {
  auth.refresh_token = res.refresh_token ?? auth.refresh_token;
  auth.tokens[scopes.join(' ')] = { access_token: res.access_token, expires_at: Date.now() + (res.expires_in - 60) * 1000 };
  await save('token.json', auth);
  return res.access_token;
}

async function pollLogin(dc, current) {
  try {
    while (Date.now() < current.expiresAt) {
      await sleep(dc.interval * 1000);
      const res = await oauth('token', { grant_type: 'urn:ietf:params:oauth:grant-type:device_code', device_code: dc.device_code });
      if (res.access_token) {
        auth.tokens = {};
        await storeToken(res, CHAT_SCOPES);
        break;
      }
      if (res.error === 'slow_down') await sleep(5000);
      else if (res.error !== 'authorization_pending') {
        loginFailure = res.error_description?.split('\r\n')[0] ?? res.error;
        break;
      }
    }
  } catch (e) {
    loginFailure = e.message;
  }
  if (login === current) login = null;
}

async function signInRequired() {
  if (!login) {
    const dc = await oauth('devicecode', { scope: scopeParam(CHAT_SCOPES) });
    if (!dc.device_code) throw new Error(`Microsoft sign-in could not start: ${dc.error_description ?? dc.error}`);
    login = { message: dc.message, expiresAt: Date.now() + dc.expires_in * 1000 };
    pollLogin(dc, login);
  }
  const previous = loginFailure ? ` (previous attempt failed: ${loginFailure})` : '';
  loginFailure = null;
  return Object.assign(new Error(`Microsoft sign-in required${previous}. Ask the user to do this: ${login.message} Then call the tool again.`), { code: 'signin' });
}

async function getToken(scopes) {
  auth ??= await load('token.json', { tokens: {} });
  const cached = auth.tokens[scopes.join(' ')];
  if (cached?.expires_at > Date.now()) return cached.access_token;
  if (auth.refresh_token) {
    const res = await oauth('token', { grant_type: 'refresh_token', refresh_token: auth.refresh_token, scope: scopeParam(scopes) });
    if (res.access_token) return storeToken(res, scopes);
    if (scopes !== CHAT_SCOPES) {
      await accessToken(CHAT_SCOPES); // throws the sign-in prompt if the session itself expired
      throw Object.assign(new Error(`channel messages are not available for this account (${res.error_description?.split('\r\n')[0] ?? res.error})`), { code: 'nochannels' });
    }
  }
  throw await signInRequired();
}

export function accessToken(scopes = CHAT_SCOPES) {
  const key = scopes.join(' ');
  if (!inflight.has(key)) inflight.set(key, getToken(scopes).finally(() => inflight.delete(key)));
  return inflight.get(key);
}

export function forgetToken(scopes) {
  if (auth) delete auth.tokens[scopes.join(' ')];
}

// ---- Graph REST

export async function graph(path, { method = 'GET', body, scopes = CHAT_SCOPES } = {}) {
  const url = path.startsWith('https://') ? path : GRAPH + path;
  for (let attempt = 0; ; attempt++) {
    const res = await fetch(url, {
      method,
      body: body && JSON.stringify(body),
      headers: { Authorization: `Bearer ${await accessToken(scopes)}`, ...(body && { 'Content-Type': 'application/json' }) },
    });
    if (attempt < 3 && (res.status === 429 || res.status === 503)) {
      await sleep(Number(res.headers.get('retry-after') || 2) * 1000);
      continue;
    }
    if (attempt === 0 && res.status === 401) {
      forgetToken(scopes);
      continue;
    }
    const data = res.status === 204 ? null : await res.json().catch(() => null);
    if (res.ok) return data;
    const message = data?.error?.message ?? res.statusText;
    throw Object.assign(new Error(`Graph ${res.status} on ${method} ${url.split('?')[0].replace(GRAPH, '')}: ${message}`), { status: res.status, graphMessage: message });
  }
}
