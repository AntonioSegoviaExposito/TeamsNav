import { createHash } from 'crypto';
import { CHANNEL_SCOPES, CHAT_SCOPES, accessToken, graph, load, save } from './graph.js';

const PAGE_CAP = 20; // Graph pages (up to 50 messages each) fetched by a single read
const FRESH_MS = 10_000; // a conversation's newest messages are re-fetched at most this often
const CHAT_TYPES = { oneOnOne: '1:1', group: 'group', meeting: 'meeting' };
const SELF_CHAT = '48:notes';
const enc = encodeURIComponent;
const time = (iso) => Date.parse(iso);
const byTime = (a, b) => time(a.at) - time(b.at) || a.id.localeCompare(b.id);

// ---- Text

const ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' };

export function htmlToText(html = '') {
  return html
    .replace(/<attachment[^>]*>\s*<\/attachment>/gi, '')
    .replace(/<img[^>]*>/gi, '[image]')
    .replace(/<at\b[^>]*>(.*?)<\/at>/gi, '@$1')
    .replace(/<li[^>]*>/gi, '- ')
    .replace(/<br\s*\/?>|<\/(p|div|li|tr|h\d)>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e) => (e[0] !== '#'
      ? ENTITIES[e.toLowerCase()] ?? m
      : String.fromCodePoint(e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : Number(e.slice(1)))))
    .replace(/\S{80,}/g, t => (/^https?:\/\//.test(t) ? t : '[long-token-omitted]'))
    .replace(/[ \t\u00a0]+/g, ' ')
    .replace(/ *\n */g, '\n')
    .replace(/\n{2,}/g, '\n')
    .trim();
}

function parseJson(text) {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

// Visible text of an Adaptive Card: TextBlock texts and FactSet "title: value" pairs.
function cardText(content) {
  const out = [];
  const walk = (node) => {
    if (Array.isArray(node)) return node.forEach(walk);
    if (!node || typeof node !== 'object') return;
    if (typeof node.text === 'string') out.push(node.text);
    if (typeof node.title === 'string' && typeof node.value === 'string') out.push(`${node.title.replace(/\s*:?\s*$/, ':')} ${node.value}`);
    Object.values(node).forEach(v => typeof v === 'object' && walk(v));
  };
  walk(parseJson(content));
  return out.join('\n').replace(/\r\n?/g, '\n').trim();
}

// Graph chatMessage -> { id, at, from, fromId, text, quote?, files }. System events and deleted messages -> null.
function normalize(m) {
  if (m.messageType !== 'message' || m.deletedDateTime) return null;
  let quote;
  const files = [];
  const cards = [];
  for (const a of m.attachments ?? []) {
    if (a.contentType === 'messageReference') {
      const c = parseJson(a.content);
      quote = { id: c?.messageId, from: c?.messageSender?.user?.displayName ?? '?', text: htmlToText(c?.messagePreview ?? '') };
    } else if (a.contentType === 'reference') files.push(`file: ${a.name}`);
    else if (a.contentType.includes('card') && cardText(a.content)) cards.push(`[card] ${cardText(a.content)}`);
    else files.push(a.contentType.includes('card') ? 'card' : a.name ?? a.contentType);
  }
  const body = m.body?.contentType === 'html' ? htmlToText(m.body.content) : (m.body?.content ?? '').trim();
  const text = [body, ...cards].filter(Boolean).join('\n');
  if (!text && !files.length) return null;
  return {
    id: m.id,
    at: m.createdDateTime,
    from: m.from?.user?.displayName ?? m.from?.application?.displayName ?? '?',
    fromId: m.from?.user?.id ?? null,
    text,
    quote,
    files,
  };
}

// ---- State: handles and names, shared by every MCP process through the cache folder

let state;
const fresh = (s) => ({ chats: {}, channels: {}, aliases: {}, roots: {}, ...s });

async function getState() {
  return (state ??= fresh(await load('state.json', {})));
}

async function persist() {
  const disk = fresh(await load('state.json', {}));
  for (const k of ['chats', 'channels', 'aliases', 'roots']) state[k] = { ...disk[k], ...state[k] };
  state.me ??= disk.me;
  await save('state.json', state);
}

// Handles are derived from the Graph id, so they are stable across processes and cache wipes.
function aliasFor(prefix, id) {
  const hash = BigInt(`0x${createHash('sha1').update(id).digest('hex').slice(0, 16)}`).toString(36).padStart(13, '0');
  for (let len = 5; ; len++) {
    const alias = prefix + hash.slice(0, len);
    if (!state.aliases[alias] || state.aliases[alias] === id) {
      state.aliases[alias] = id;
      return alias;
    }
  }
}

const chatMeta = (chatId) => (state.chats[chatId] ??= { alias: aliasFor('C', chatId) });

function channelMeta(teamId, channelId) {
  const key = `${teamId}|${channelId}`;
  return (state.channels[key] ??= { alias: aliasFor('K', key), teamId, channelId });
}

export async function me() {
  await getState();
  if (!state.me) {
    const u = await graph('/me?$select=id,displayName');
    state.me = { id: u.id, name: u.displayName };
  }
  return state.me;
}

let channelsOk;
async function channelScopes() {
  if (channelsOk === undefined) {
    try {
      await accessToken(CHANNEL_SCOPES);
      channelsOk = true;
    } catch (e) {
      if (e.code !== 'nochannels') throw e;
      channelsOk = false;
    }
  }
  return channelsOk ? CHANNEL_SCOPES : CHAT_SCOPES;
}

async function mapLimit(items, limit, fn) {
  let next = 0;
  const worker = async () => {
    while (next < items.length) await fn(items[next++]);
  };
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
}

// ---- Names

function chatName(chat, members, myId) {
  if (chat.topic) return chat.topic;
  const others = members.filter(m => m.userId !== myId).map(m => m.displayName).filter(Boolean);
  if (!others.length) return chat.chatType === 'oneOnOne' ? 'me' : '(no name)';
  return others.length > 3 ? `${others.slice(0, 3).join(', ')} +${others.length - 3}` : others.join(', ');
}

// Names chats from Graph chat objects; members are fetched only for unnamed chats not seen before.
async function describeChats(chats) {
  const { id: myId } = await me();
  await mapLimit(chats.filter(c => !c.topic && !c.members && !state.chats[c.id]?.name), 8, async (c) => {
    c.members = (await graph(`/chats/${enc(c.id)}/members`)).value;
  });
  for (const c of chats) {
    const meta = chatMeta(c.id);
    meta.type = CHAT_TYPES[c.chatType] ?? c.chatType;
    if (c.topic || c.members) meta.name = chatName(c, c.members ?? [], myId);
  }
}

async function nameChats(chatIds) {
  const unknown = chatIds.filter(id => !state.chats[id]?.name);
  const found = [];
  await mapLimit(unknown, 8, async (id) => {
    if (id === SELF_CHAT) Object.assign(chatMeta(id), { name: 'notes to self', type: '1:1' });
    else found.push(await graph(`/chats/${enc(id)}?$expand=members`).catch(() => ({ id, chatType: '?', topic: '(unavailable chat)' })));
  });
  await describeChats(found);
}

async function nameChannels(channels) {
  await mapLimit(channels.filter(c => !c.name), 8, async (ch) => {
    const opts = { scopes: CHANNEL_SCOPES };
    const [team, channel] = await Promise.all([
      graph(`/teams/${ch.teamId}?$select=displayName`, opts).catch(() => ({ displayName: '?' })),
      graph(`/teams/${ch.teamId}/channels/${enc(ch.channelId)}?$select=displayName`, opts).catch(() => ({ displayName: '?' })),
    ]);
    ch.name = `${team.displayName} › ${channel.displayName}`;
  });
}

// ---- Chats

export async function listChats({ query, limit }) {
  await getState();
  const { id: myId } = await me();
  const q = query?.toLowerCase();
  const rows = [];
  let url = '/me/chats?$top=50&$expand=lastMessagePreview&$orderby=lastMessagePreview/createdDateTime desc';
  for (let page = 0; url && rows.length < limit && page < 10; page++) {
    const { value, '@odata.nextLink': next } = await graph(url);
    const visible = value.filter(c => !c.viewpoint?.isHidden);
    await describeChats(visible);
    for (const c of visible) {
      const meta = state.chats[c.id];
      if (q && !meta.name.toLowerCase().includes(q)) continue;
      const last = c.lastMessagePreview;
      const read = c.viewpoint?.lastMessageReadDateTime;
      rows.push({
        handle: meta.alias,
        type: meta.type,
        name: meta.name,
        at: last?.createdDateTime ?? c.lastUpdatedDateTime,
        preview: last && `${last.from?.user?.displayName ?? '?'}: ${htmlToText(last.body?.content)}`,
        unread: !!last && last.from?.user?.id !== myId && (!read || time(read) < time(last.createdDateTime)),
      });
    }
    url = next;
  }
  await persist();
  return rows.slice(0, limit);
}

// ---- Conversations: a chat, or a channel thread (root post + replies). Both are read oldest-first.

const historyFile = (conv) => `messages/${createHash('sha1').update(conv.key).digest('hex').slice(0, 16)}.json`;
const chatPage = (chatId, before) =>
  `/chats/${enc(chatId)}/messages?$top=50&$orderby=createdDateTime desc${before ? `&$filter=createdDateTime lt ${before}` : ''}`;

function merge(hist, raw) {
  const byId = new Map(hist.messages.map(m => [m.id, m]));
  for (const r of raw) {
    const m = normalize(r);
    if (m) byId.set(m.id, m);
    else byId.delete(r.id);
  }
  hist.messages = [...byId.values()].sort(byTime);
}

// A chat history is one contiguous span [from, to] of the chat, grown backwards on demand.
async function refreshChat(conv, hist) {
  const raw = [];
  let before;
  let joined = !hist.to; // with nothing cached, the newest page is enough
  for (let page = 0; page < PAGE_CAP; page++) {
    const { value, '@odata.nextLink': next } = await graph(chatPage(conv.id, before));
    raw.push(...value);
    const oldest = value.at(-1)?.createdDateTime;
    if (!next || !oldest) {
      hist.complete = joined = true;
      break;
    }
    if (joined || time(oldest) <= time(hist.to)) {
      joined = true;
      break;
    }
    before = oldest;
  }
  if (!joined) Object.assign(hist, { messages: [], from: null, complete: false }); // too many new messages: restart the span
  merge(hist, raw);
  if (raw.length) {
    hist.to = raw[0].createdDateTime;
    hist.from ??= raw.at(-1).createdDateTime;
  }
}

async function extendChat(conv, hist) {
  const { value, '@odata.nextLink': next } = await graph(chatPage(conv.id, hist.from));
  merge(hist, value);
  if (value.length) hist.from = value.at(-1).createdDateTime;
  if (!next || !value.length) hist.complete = true;
}

async function refreshThread(conv, hist) {
  const opts = { scopes: CHANNEL_SCOPES };
  const base = `/teams/${conv.teamId}/channels/${enc(conv.channelId)}/messages/${conv.rootId}`;
  const raw = [await graph(base, opts)];
  let url = `${base}/replies?$top=50`;
  for (let page = 0; url && page < PAGE_CAP; page++) {
    const { value, '@odata.nextLink': next } = await graph(url, opts);
    raw.push(...value);
    url = next;
  }
  hist.messages = [];
  merge(hist, raw);
  hist.complete = true;
}

async function history(conv, force = false) {
  const hist = await load(historyFile(conv), { messages: [] });
  if (force || Date.now() - (hist.checkedAt ?? 0) >= FRESH_MS) {
    await (conv.kind === 'chat' ? refreshChat : refreshThread)(conv, hist);
    hist.checkedAt = Date.now();
  }
  return hist;
}

export async function latest(conv, count) {
  const hist = await history(conv);
  for (let p = 0; hist.messages.length < count && !hist.complete && p < PAGE_CAP; p++) await extendChat(conv, hist);
  await save(historyFile(conv), hist);
  return { messages: hist.messages.slice(-count), older: hist.messages.length > count || !hist.complete };
}

export async function around(conv, id, before, after) {
  const hist = await history(conv);
  const index = () => hist.messages.findIndex(m => m.id === id);
  let pages = 0;
  while (index() < before && !hist.complete && pages++ < PAGE_CAP) await extendChat(conv, hist);
  await save(historyFile(conv), hist);
  const i = index();
  if (i < 0 && hist.complete) throw new Error(`message ${id} is not a text message of this conversation (deleted, a system event, or a wrong handle)`);
  if (i < 0) throw new Error(`message ${id} is older than the history loaded so far (back to ${hist.from}); call again to keep loading`);
  const start = Math.max(0, i - before);
  return {
    messages: hist.messages.slice(start, i + after + 1),
    older: start > 0 || !hist.complete,
    newer: i + after + 1 < hist.messages.length,
  };
}

// Messages newer than `iso`, re-fetching the newest page right now (used by chat mode).
export async function since(conv, iso) {
  const hist = await history(conv, true);
  await save(historyFile(conv), hist);
  return hist.messages.filter(m => time(m.at) > time(iso));
}

export async function channelPosts(channel, count) {
  const { value } = await graph(`/teams/${channel.teamId}/channels/${enc(channel.channelId)}/messages?$top=${Math.min(count, 50)}`, { scopes: CHANNEL_SCOPES });
  return value.map(normalize).filter(Boolean).sort(byTime).slice(-count);
}

// Graph only serves thread roots on the /chats route; for a reply, its 400 error names the root:
// "... retrieve replies via /chats(<channel>)/messages(<root>)/replies(<id>)".
async function threadRoot(channel, id) {
  const key = `${channel.channelId}|${id}`;
  if (!state.roots[key]) {
    let root = id;
    try {
      await graph(`/chats/${enc(channel.channelId)}/messages/${id}`, { scopes: CHANNEL_SCOPES });
    } catch (e) {
      const m = e.status === 400 && /\/messages\((\d+)\)\/replies\(/.exec(e.graphMessage ?? '');
      if (!m) throw e;
      root = m[1];
    }
    state.roots[key] = root;
  }
  return state.roots[key];
}

async function chatConversation(chatId) {
  await nameChats([chatId]);
  const meta = chatMeta(chatId);
  return { kind: 'chat', id: chatId, key: chatId, handle: meta.alias, title: `${meta.type ?? 'chat'} · ${meta.name}` };
}

// target: "C…" / "K…" handle, "<handle>/<messageId>", a raw Graph chat id, or a chat name.
export async function resolveTarget(target) {
  await getState();
  const m = /^(.*)\/(\d+)$/.exec(target.trim());
  const [part, messageId] = m ? [m[1], m[2]] : [target.trim(), null];
  let id = /^[CK][0-9a-z]{5,}$/.test(part) && state.aliases[part];
  if (!id && /^[CK][0-9a-z]{5,}$/.test(part)) {
    state = fresh(await load('state.json', {})); // the handle may come from another MCP process
    id = state.aliases[part];
  }
  let conv;
  if (id && part[0] === 'C') conv = await chatConversation(id);
  else if (id) {
    const channel = state.channels[id];
    await nameChannels([channel]);
    const base = { teamId: channel.teamId, channelId: channel.channelId, handle: channel.alias, title: `channel · ${channel.name}` };
    const rootId = messageId && (await threadRoot(channel, messageId));
    conv = rootId ? { ...base, kind: 'thread', rootId, key: `${channel.channelId}|${rootId}` } : { ...base, kind: 'channel' };
  } else if (/^\d+:.+/.test(part)) conv = await chatConversation(part);
  else {
    const matches = await listChats({ query: part, limit: 10 });
    if (matches.length !== 1) {
      const list = matches.map(r => `${r.handle}  ${r.type}  ${r.name}`).join('\n');
      if (matches.length) throw new Error(`"${part}" matches ${matches.length} chats; use a handle:\n${list}`);
      const handleLike = /^[CK][0-9a-z]{5,}$/.test(part) ? `unknown handle "${part}" (handles come from teams_chats, teams_search or teams_read) and ` : '';
      throw new Error(`${handleLike}no chat is named like "${part}" among the 500 most recent chats; try teams_search`);
    }
    conv = await chatConversation(state.aliases[matches[0].handle]);
  }
  await persist();
  return { ...conv, messageId };
}

export async function send(conv, text) {
  if (conv.kind !== 'chat') throw new Error('sending is only supported in chats (C handles), not in channels');
  const html = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\r?\n/g, '<br>');
  const sent = await graph(`/chats/${enc(conv.id)}/messages`, { method: 'POST', body: { body: { contentType: 'html', content: html } } });
  const hist = await load(historyFile(conv), null);
  if (hist) await save(historyFile(conv), { ...hist, checkedAt: 0 }); // next read re-fetches the newest messages
  return normalize(sent);
}

// ---- Search

export async function search(query, limit, offset) {
  await getState();
  const res = await graph('/search/query', {
    method: 'POST',
    scopes: await channelScopes(),
    body: { requests: [{ entityTypes: ['chatMessage'], query: { queryString: query }, from: offset, size: limit }] },
  });
  const box = res.value?.[0]?.hitsContainers?.[0] ?? {};
  const hits = (box.hits ?? []).map(({ resource: r, summary }) => {
    const ch = r.channelIdentity?.channelId && r.channelIdentity;
    return {
      place: ch ? channelMeta(ch.teamId, ch.channelId) : chatMeta(r.chatId),
      chatId: ch ? null : r.chatId,
      id: r.id,
      at: r.createdDateTime,
      from: r.from?.emailAddress?.name ?? '?',
      text: htmlToText((summary ?? '').replace(/<ddd\/>/g, '…')),
    };
  });
  await nameChats([...new Set(hits.filter(h => h.chatId).map(h => h.chatId))]);
  await nameChannels([...new Set(hits.filter(h => !h.chatId).map(h => h.place))]);
  await persist();
  return { total: box.total ?? 0, more: !!box.moreResultsAvailable, hits };
}
