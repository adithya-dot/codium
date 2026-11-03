# Samsung Notes MCP connector

An MCP server that gives Claude (or any MCP client) access to your Samsung Notes
through Microsoft Graph.

Samsung Notes has no public API. The supported way to get notes off the phone
is the app's **Sync to Microsoft OneNote** option, which pushes them into your
Microsoft account (they appear in the OneNote feed in Outlook). This server
reads them from there.

## Tools

| Tool | What it does |
| --- | --- |
| `find_note_folders` | Search your mailbox (including hidden folders) for folders holding synced notes. `show_all=true` lists every folder. |
| `list_notes` | List notes in a folder, newest first. |
| `get_note` | Full text of one note. |
| `search_notes` | Full-text search in a notes folder. |
| `list_onenote_sections` | List sections in your OneNote notebooks. |
| `list_onenote_pages` | List OneNote pages, by section or title. |
| `get_onenote_page` | Text of a OneNote page. |
| `create_onenote_page` | Create a OneNote page. |

Synced Samsung Notes are **read-only**: Samsung only syncs one way, so edits
must be made in the phone app. `create_onenote_page` writes to OneNote, not to
Samsung Notes.

> Microsoft does not document which mailbox folder holds feed notes, so run
> `find_note_folders` first (try `show_all=true` if nothing matches) and use the
> folder that contains your notes.

## Setup

1. **Turn on sync on your phone:** Samsung Notes → ☰ → Settings → *Sync to
   Microsoft OneNote*, sign in, and pick the folders to sync.
2. **Register an Azure app** (free, one-time):
   - Go to <https://entra.microsoft.com> → *App registrations* → *New registration*.
   - Account types: *Personal Microsoft accounts only* (or *any org + personal*
     if you use a work account).
   - Under *Authentication*, enable **Allow public client flows**.
   - Under *API permissions*, add delegated Microsoft Graph permissions
     `Mail.Read`, `Notes.ReadWrite`, `User.Read`.
   - Copy the **Application (client) ID**.
3. **Install and sign in:**
   ```bash
   cd samsung-notes-mcp
   pip install -r requirements.txt
   export SAMSUNG_NOTES_CLIENT_ID=<your-client-id>
   python server.py login   # follow the device-code prompt
   ```
   The token is cached at `~/.samsung-notes-mcp/token_cache.json` (mode 600).
   For a work/school account, also set
   `SAMSUNG_NOTES_AUTHORITY=https://login.microsoftonline.com/organizations`.
4. **Add it to Claude Code:**
   ```bash
   claude mcp add samsung-notes \
     -e SAMSUNG_NOTES_CLIENT_ID=<your-client-id> \
     -- python /absolute/path/to/samsung-notes-mcp/server.py
   ```
   Or for Claude Desktop, in `claude_desktop_config.json`:
   ```json
   {
     "mcpServers": {
       "samsung-notes": {
         "command": "python",
         "args": ["/absolute/path/to/samsung-notes-mcp/server.py"],
         "env": { "SAMSUNG_NOTES_CLIENT_ID": "<your-client-id>" }
       }
     }
   }
   ```

## Environment variables

| Variable | Default |
| --- | --- |
| `SAMSUNG_NOTES_CLIENT_ID` | required |
| `SAMSUNG_NOTES_AUTHORITY` | `https://login.microsoftonline.com/consumers` |
| `SAMSUNG_NOTES_TOKEN_CACHE` | `~/.samsung-notes-mcp/token_cache.json` |
