#!/usr/bin/env node
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import { sleep } from './graph.js';
import { around, channelPosts, latest, listChats, me, resolveTarget, search, send, since } from './teams.js';

const pad = (n) => String(n).padStart(2, '0');
const when = (iso) => {
  const d = new Date(iso);
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
};
const clip = (text, max) => (text.length > max ? `${text.slice(0, max)}…` : text);
const MAX_TEXT = 600;

// "<handle>/<id>  <date>  <sender>: <text>"; continuation lines indented.
function line(handle, m, myId, { full = false, mark = '' } = {}) {
  let text = m.text;
  if (m.quote) text = `(re ${m.quote.id ? `${handle}/${m.quote.id} ` : ''}${m.quote.from}: "${clip(m.quote.text.replace(/\n/g, ' '), 60)}") ${text}`;
  if (m.files.length) text = `${text} [${m.files.join(', ')}]`.trim();
  if (!full && text.length > MAX_TEXT) text = `${text.slice(0, MAX_TEXT)}… [+${text.length - MAX_TEXT} chars]`;
  const who = m.fromId && m.fromId === myId ? `${m.from} (me)` : m.from;
  return `${mark}${handle}/${m.id}  ${when(m.at)}  ${who}: ${text.replace(/\n/g, '\n    ')}`;
}

const server = new McpServer(
  { name: 'teams-nav', version: '1.0.0' },
  {
    instructions: 'Microsoft Teams through Microsoft Graph, acting as the signed-in user. Chats have handles like Cab12x, channels like Kab12x, and a message is <handle>/<messageId>. Typical flow: find with teams_search or teams_chats, read the conversation around a message with teams_read, write with teams_send or teams_chat_mode. On first use a tool returns a Microsoft sign-in code: show it to the user and call the tool again once they finish.',
  },
);

function tool(name, config, run) {
  server.registerTool(name, config, async (args, extra) => {
    try {
      return { content: [{ type: 'text', text: await run(args, extra) }] };
    } catch (e) {
      return { content: [{ type: 'text', text: `ERR ${e.message}` }], isError: true };
    }
  });
}

tool(
  'teams_chats',
  {
    title: 'List Teams chats',
    description: 'List the signed-in user\'s Teams chats (1:1, group and meeting chats), most recent first. Each row: chat handle (C…), type, name, time and sender of the last message, and [unread] when the user has not read it. Pass `query` to keep only chats whose name contains it (for 1:1 chats the name is the other person). Use the handle with teams_read, teams_send or teams_chat_mode.',
    inputSchema: {
      query: z.string().optional().describe('Text the chat name must contain, e.g. a person or a group topic.'),
      limit: z.number().int().min(1).max(100).optional().describe('Rows to return (default 20).'),
    },
  },
  async ({ query, limit = 20 }) => {
    const rows = await listChats({ query, limit });
    const head = `CHATS  ·  ${rows.length} ${query ? `matching "${query}"` : 'most recent'}`;
    const body = rows.map(r => `${r.handle}  ${r.type.padEnd(7)}  ${r.name}  ·  ${when(r.at)}  ${clip((r.preview ?? '').replace(/\n/g, ' '), 80)}${r.unread ? '  [unread]' : ''}`);
    return [head, '', ...body].join('\n') + (rows.length ? '' : 'No chats.');
  },
);

tool(
  'teams_search',
  {
    title: 'Search Teams messages',
    description: 'Search the Teams messages the signed-in user can see (chats and channels) by keywords. Uses Microsoft Search: words, "exact phrases", OR. Results are ranked by relevance; each line starts with a message handle (C…/id in chats, K…/id in channels), then date, place, sender and a snippet. Pass a handle to teams_read to see what was said before and after it. Page with `offset`.',
    inputSchema: {
      query: z.string().min(1).describe('Keywords to search for.'),
      limit: z.number().int().min(1).max(100).optional().describe('Results to return (default 20).'),
      offset: z.number().int().min(0).optional().describe('Results to skip, for the next page (default 0).'),
    },
  },
  async ({ query, limit = 20, offset = 0 }) => {
    const { total, more, hits } = await search(query, limit, offset);
    const head = `SEARCH "${query}"  ·  ${hits.length ? `${offset + 1}-${offset + hits.length}` : 0} of ${total}${more ? `  ·  next page: offset ${offset + hits.length}` : ''}`;
    const body = hits.map(h => `${h.place.alias}/${h.id}  ${when(h.at)}  ${h.place.name}  ·  ${h.from}: ${h.text.replace(/\n/g, ' ')}`);
    return [head, '', ...body].join('\n');
  },
);

tool(
  'teams_read',
  {
    title: 'Read a Teams conversation',
    description: [
      'Read a Teams conversation, oldest message first.',
      '- target = chat handle (C…), chat name or raw chat id: its latest `before` messages (default 20).',
      '- target = message handle (C…/id or K…/id, from teams_search or a previous read): `before` messages before it (default 10), the message itself marked with » and shown in full, and `after` messages after it (default 10). Channel messages are read inside their thread (root post + replies).',
      '- target = channel handle (K…): its latest posts; open a post with its handle to read the thread.',
      'Lines are "<handle>/<id>  <date time>  <sender>: <text>". Any line handle is a valid new target: e.g. the first line with after 0 scrolls back. Messages longer than 600 characters are cut, except the target. "(re …)" is a quoted reply and "(me)" the signed-in user.',
    ].join('\n'),
    inputSchema: {
      target: z.string().min(1).describe('Chat handle, message handle, channel handle, chat name or raw chat id.'),
      before: z.number().int().min(0).max(100).optional().describe('Messages before the target message (default 10), or latest messages for a chat or channel (default 20).'),
      after: z.number().int().min(0).max(100).optional().describe('Messages after the target message (default 10).'),
    },
  },
  async ({ target, before, after = 10 }) => {
    const conv = await resolveTarget(target);
    const { id: myId } = await me();
    const head = `${conv.kind.toUpperCase()} ${conv.handle}  ·  ${conv.title}`;
    if (conv.kind === 'channel') {
      const posts = await channelPosts(conv, before ?? 20);
      return [`${head}  ·  latest ${posts.length} posts`, '', ...posts.map(m => line(conv.handle, m, myId))].join('\n');
    }
    if (!conv.messageId) {
      const { messages, older } = await latest(conv, before ?? 20);
      const lines = messages.map(m => line(conv.handle, m, myId));
      if (older && messages.length) lines.unshift(`… older messages: target "${conv.handle}/${messages[0].id}" with after 0`);
      return [head, '', ...(lines.length ? lines : ['No messages.'])].join('\n');
    }
    const { messages, older, newer } = await around(conv, conv.messageId, before ?? 10, after);
    const lines = messages.map(m => line(conv.handle, m, myId, m.id === conv.messageId ? { full: true, mark: '» ' } : { mark: '  ' }));
    if (older) lines.unshift(`… older messages: target "${conv.handle}/${messages[0].id}" with after 0`);
    if (newer) lines.push(`… newer messages: target "${conv.handle}/${messages.at(-1).id}" with before 0`);
    return [head, '', ...lines].join('\n');
  },
);

tool(
  'teams_send',
  {
    title: 'Send a Teams message',
    description: 'Send a message to a Teams chat as the signed-in user. Only send what the user asked to send. `chat` accepts a chat handle, a message handle (sends to that message\'s chat), a chat name or a raw chat id. Line breaks are kept. Returns the handle of the sent message. Channels are read-only.',
    inputSchema: {
      chat: z.string().min(1).describe('Chat handle (C…), message handle, chat name or raw chat id.'),
      text: z.string().min(1).describe('Message text.'),
    },
  },
  async ({ chat, text }) => {
    const conv = await resolveTarget(chat);
    const sent = await send(conv, text);
    return `SENT  ·  ${conv.title}\n${line(conv.handle, sent, (await me()).id, { full: true })}`;
  },
);

// ---- Chat mode: the server keeps the cursor and closes idle sessions, so the agent only loops on wait/reply.

const IDLE_MS = 10 * 60_000;
const POLL_MS = 5000;
const BOT = '🤖';
const sessions = new Map(); // chat id -> { conv, cursor, idleSince, timer }
const withBot = (text) => (text.startsWith(BOT) ? text : `${BOT} ${text}`);

async function close(session, text) {
  if (sessions.get(session.conv.id) !== session) return;
  sessions.delete(session.conv.id);
  clearTimeout(session.timer);
  await send(session.conv, withBot(text)).catch(() => {});
}

function arm(session) {
  clearTimeout(session.timer);
  session.idleSince = Date.now();
  session.timer = setTimeout(() => {
    session.reason = 'turned off after 10 minutes without messages';
    close(session, 'Chat mode off after 10 minutes without messages.');
  }, IDLE_MS);
}

async function waitFor(session, seconds, extra) {
  const { id: myId } = await me();
  const started = Date.now();
  const progressToken = extra?._meta?.progressToken;
  for (;;) {
    if (sessions.get(session.conv.id) !== session) return null;
    const incoming = (await since(session.conv, session.cursor)).filter(m => m.fromId !== myId);
    if (incoming.length) {
      session.cursor = incoming.at(-1).at;
      arm(session);
      return incoming;
    }
    const elapsed = Date.now() - started;
    if (elapsed >= seconds * 1000) return [];
    if (progressToken !== undefined) {
      await extra.sendNotification({ method: 'notifications/progress', params: { progressToken, progress: Math.round(elapsed / 1000), total: seconds, message: 'waiting for messages' } });
    }
    await sleep(Math.min(POLL_MS, seconds * 1000 - elapsed));
  }
}

tool(
  'teams_chat_mode',
  {
    title: 'Chat on the user\'s behalf',
    description: [
      'Hold a conversation in a Teams chat on the user\'s behalf. Use it only when the user explicitly asks you to talk in a chat for them. Every message you post is prefixed with 🤖.',
      '- start: posts `text` announcing that an AI agent is answering from the user\'s account (write it in the chat\'s language; default is an English notice), then waits.',
      '- wait: waits for new messages from other people.',
      '- reply: posts `text`, then waits.',
      '- stop: posts `text` (default "Chat mode off.") and ends chat mode.',
      'Waits return as soon as someone else writes, or after `seconds` with nothing new: then call wait again. Messages that arrive while you are replying are never lost. Chat mode turns itself off, with a notice in the chat, after 10 minutes without messages from others. Never present yourself as the user, never share secrets, and ask the user before agreeing to anything on their behalf.',
    ].join('\n'),
    inputSchema: {
      action: z.enum(['start', 'wait', 'reply', 'stop']),
      chat: z.string().min(1).describe('Chat handle (C…), message handle, chat name or raw chat id.'),
      text: z.string().optional().describe('Message for start, reply (required) and stop.'),
      seconds: z.number().int().min(5).max(600).optional().describe('Maximum wait for new messages (default 50).'),
    },
  },
  async ({ action, chat, text, seconds = 50 }, extra) => {
    const conv = await resolveTarget(chat);
    if (conv.kind !== 'chat') throw new Error('chat mode only works in chats (C handles)');
    const head = `CHAT MODE ${conv.handle}  ·  ${conv.title}`;
    let session = sessions.get(conv.id);
    if (action === 'start' && !session) {
      const sent = await send(conv, withBot(text ?? `Chat mode on: an AI agent is replying from ${(await me()).name}'s account. It turns off after 10 minutes without messages.`));
      session = { conv, cursor: sent.at };
      sessions.set(conv.id, session);
      arm(session);
    }
    if (!session) throw new Error(`chat mode is not on in ${conv.handle}; use action "start" first`);
    if (action === 'stop') {
      await close(session, text ?? 'Chat mode off.');
      return `${head}  ·  off`;
    }
    if (action === 'reply') {
      if (!text) throw new Error('reply needs `text`');
      await send(conv, withBot(text));
      arm(session);
    }
    const incoming = await waitFor(session, seconds, extra);
    if (!incoming) return `${head}  ·  off (${session.reason ?? 'stopped'})`;
    if (!incoming.length) {
      const idle = Math.floor((Date.now() - session.idleSince) / 60_000);
      return `${head}  ·  on  ·  no new messages in ${seconds} s  ·  idle ${idle} min (turns off at 10)`;
    }
    const { id: myId } = await me();
    return [`${head}  ·  on  ·  ${incoming.length} new`, '', ...incoming.map(m => line(conv.handle, m, myId, { full: true }))].join('\n');
  },
);

// Leaving: tell every chat still in chat mode that the agent is gone.
async function shutdown() {
  await Promise.race([Promise.all([...sessions.values()].map(s => close(s, 'Chat mode off.'))), sleep(3000)]);
  process.exit(0);
}
process.stdin.on('close', shutdown);
process.on('SIGINT', shutdown);
process.on('SIGTERM', shutdown);

await server.connect(new StdioServerTransport());
