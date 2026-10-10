# TeamsNav

MCP server that lets an agent search, read and write Microsoft Teams messages through Microsoft Graph, acting as the signed-in user. It finds messages by keyword, shows the conversation before and after any message, lists chats, sends messages and can hold a conversation on the user's behalf ("chat mode"). Tool descriptions are complete usage manuals, so agents need no extra instructions.

## Configuration

| Variable | Default | Meaning |
|----------|---------|---------|
| `TENANT` | `organizations` | Microsoft Entra tenant (domain or id), e.g. `contoso.com` |
| `CLIENT_ID` | `14d82eec-204b-4c2f-b7e8-296a70dab67e` | Public client for the device-code sign-in. Default is Microsoft Graph Command Line Tools; no secret. |

opencode (`~/.config/opencode/opencode.json`):

```json
{
  "mcp": {
    "teams-nav": {
      "type": "local",
      "command": ["pnpm", "dlx", "github:AntonioSegoviaExposito/TeamsNav"],
      "environment": { "TENANT": "contoso.com", "CLIENT_ID": "14d82eec-204b-4c2f-b7e8-296a70dab67e" },
      "enabled": true
    }
  }
}
```

Clients using the `mcpServers` format (Claude Desktop, Cursor, …):

```json
{
  "mcpServers": {
    "teams-nav": {
      "command": "pnpm",
      "args": ["dlx", "github:AntonioSegoviaExposito/TeamsNav"],
      "env": { "TENANT": "contoso.com", "CLIENT_ID": "14d82eec-204b-4c2f-b7e8-296a70dab67e" }
    }
  }
}
```

### Sign-in and permissions

The first tool call returns a Microsoft device-code message (URL + code). The agent shows it to the user, who signs in once in a browser. After that the refresh token is renewed silently.

| Permission | Used for |
|------------|----------|
| `Chat.ReadWrite`, `User.Read` | Chats: list, read, search, send, chat mode |
| `ChannelMessage.Read.All`, `Team.ReadBasic.All`, `Channel.ReadBasic.All` | Channel posts in search results and their threads (optional: without them, channel messages are reported as unavailable) |

## Tools

| Tool | Does |
|------|------|
| `teams_chats` | Recent chats (1:1, group, meeting) with handle, type, last message and `[unread]`; filter by name |
| `teams_search` | Keyword search over chat and channel messages, each result with its message handle |
| `teams_read` | Latest messages of a chat or channel, or the messages before and after a given message (channel messages are read inside their thread) |
| `teams_send` | Send a message to a chat. Prefixed with 🤖 unless `belikehuman` is true. `text` and `message` are aliases; passing both is an error |
| `teams_chat_mode` | `start` / `wait` / `reply` / `stop` on the user's behalf. Same 🤖 prefix and `text`/`message` rules as `teams_send` |

Handles are short and stable: chats `C` + 5 characters, channels `K` + 5 characters, messages `<handle>/<messageId>`. Every line an agent reads starts with a handle it can pass back to `teams_read`, `teams_send` or `teams_chat_mode`.

```
CHAT C25nf0  ·  group · API StoreFront Checkout PB

… older messages: target "C25nf0/1791185500169" with after 0
  C25nf0/1791185500169  2026-10-05 09:31  Ana López: there is no storefront endpoint with that data yet
» C25nf0/1791185655198  2026-10-05 09:34  Juan Pérez: then we keep using the current one?
  C25nf0/1791186190962  2026-10-05 09:43  Ana López: (re C25nf0/1791185655198 Juan Pérez: "then we keep using the current…") yes, for now
```

Chat mode keeps its cursor in the server, so messages that arrive while the agent replies are never lost. It turns itself off with a notice in the chat after 10 minutes without messages from others, or when the server stops.

## Cache

Everything lives under `<os temp dir>/teams-nav/<hash of tenant + client>/`:

- `token.json` (mode 600): refresh token and access tokens.
- `state.json`: handles, chat and channel names.
- `messages/`: one file per chat or channel thread with the messages already fetched.

A chat's cached history is a contiguous span that grows backwards as older messages are requested. Its newest messages are re-fetched when older than 10 seconds, and right after sending. Deleting the folder only costs a new sign-in. Handles do not change, because they are derived from the Teams ids.

## Limits

- Channels are read-only.
- `teams_chats` filtering looks at the 500 most recent chats.
- Reading around a very old message pages back through the chat 1,000 messages per call. When the message is further back, the tool says so and the next call continues from there.
