"""Samsung Notes MCP connector.

Samsung Notes has no public API. When "Sync to Microsoft OneNote" is turned on
in the Samsung Notes app, notes are pushed to the user's Microsoft account and
appear in the OneNote feed / Outlook. This server reaches them through the
Microsoft Graph API:

* Samsung Notes synced to the feed live in the Exchange mailbox, not in OneNote
  notebooks. Microsoft does not document the exact folder, so `find_note_folders`
  searches all mail folders (including hidden ones) and `list_notes` reads the
  one you pick. Samsung says these notes are read-only outside the phone app.
* Regular OneNote notebooks/pages are also exposed (read + create page), which
  covers notes you move or copy into OneNote.

Auth: MSAL device-code flow against your own Azure app registration
(public client). Tokens are cached in ~/.samsung-notes-mcp/token_cache.json.
"""
from __future__ import annotations

import html
import os
import re
import sys
from pathlib import Path
from typing import Any

import httpx
import msal
from mcp.server.fastmcp import FastMCP

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = ["Mail.Read", "Notes.ReadWrite", "User.Read"]
CLIENT_ID = os.environ.get("SAMSUNG_NOTES_CLIENT_ID", "")
AUTHORITY = os.environ.get(
    "SAMSUNG_NOTES_AUTHORITY", "https://login.microsoftonline.com/consumers"
)
CACHE_PATH = Path(
    os.environ.get(
        "SAMSUNG_NOTES_TOKEN_CACHE",
        Path.home() / ".samsung-notes-mcp" / "token_cache.json",
    )
)
# Folder-name hints used to spot where synced notes landed.
NOTE_FOLDER_HINTS = re.compile(r"note|samsung|sticky", re.IGNORECASE)

mcp = FastMCP("samsung-notes")

# ---------------------------------------------------------------- auth


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if CACHE_PATH.exists():
        cache.deserialize(CACHE_PATH.read_text())
    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
        CACHE_PATH.write_text(cache.serialize())
        os.chmod(CACHE_PATH, 0o600)


def _app(cache: msal.SerializableTokenCache) -> msal.PublicClientApplication:
    if not CLIENT_ID:
        raise RuntimeError(
            "SAMSUNG_NOTES_CLIENT_ID is not set. See README.md for how to "
            "register an Azure app and get a client ID."
        )
    return msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)


def get_token(interactive: bool = True) -> str:
    cache = _load_cache()
    app = _app(cache)
    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if not result and interactive:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Could not start device flow: {flow}")
        # stdout is the MCP channel, so prompts go to stderr.
        print(flow["message"], file=sys.stderr, flush=True)
        result = app.acquire_token_by_device_flow(flow)
    _save_cache(cache)
    if not result or "access_token" not in result:
        raise RuntimeError(
            "Not signed in. Run `python server.py login` in a terminal first."
            + (f" ({result.get('error_description')})" if result else "")
        )
    return result["access_token"]


def graph(method: str, path: str, **kw: Any) -> httpx.Response:
    url = path if path.startswith("http") else GRAPH + path
    headers = {"Authorization": f"Bearer {get_token(interactive=False)}"}
    headers.update(kw.pop("headers", {}))
    r = httpx.request(method, url, headers=headers, timeout=30, **kw)
    if r.status_code >= 400:
        raise RuntimeError(f"Graph {method} {path} -> {r.status_code}: {r.text[:500]}")
    return r


def _strip_html(s: str) -> str:
    s = re.sub(r"(?is)<(script|style|head).*?</\1>", "", s or "")
    s = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</li>", "\n", s)
    s = re.sub(r"<[^>]+>", "", s)
    return re.sub(r"\n{3,}", "\n\n", html.unescape(s)).strip()


def _all_mail_folders() -> list[dict]:
    """Breadth-first walk of every mail folder, hidden ones included."""
    out: list[dict] = []
    queue = ["/me/mailFolders?includeHiddenFolders=true&$top=100"]
    while queue:
        url = queue.pop(0)
        while url:
            data = graph("GET", url).json()
            for f in data.get("value", []):
                out.append(f)
                if f.get("childFolderCount"):
                    queue.append(
                        f"/me/mailFolders/{f['id']}/childFolders"
                        "?includeHiddenFolders=true&$top=100"
                    )
            url = data.get("@odata.nextLink")
    return out


# ---------------------------------------------------------------- tools: synced notes


@mcp.tool()
def find_note_folders(show_all: bool = False) -> list[dict]:
    """Find mailbox folders that may hold Samsung Notes synced via the OneNote feed.

    Returns folders whose name mentions notes/Samsung/sticky (or every folder
    when show_all=True), with item counts. Pass the folder id to list_notes.
    """
    rows = [
        {
            "id": f["id"],
            "name": f.get("displayName"),
            "items": f.get("totalItemCount"),
            "hidden": f.get("isHidden", False),
        }
        for f in _all_mail_folders()
        if show_all or NOTE_FOLDER_HINTS.search(f.get("displayName") or "")
    ]
    return rows


@mcp.tool()
def list_notes(folder_id: str, top: int = 25, skip: int = 0) -> list[dict]:
    """List notes in a mailbox folder found by find_note_folders (newest first)."""
    data = graph(
        "GET",
        f"/me/mailFolders/{folder_id}/messages",
        params={
            "$top": min(max(top, 1), 100),
            "$skip": max(skip, 0),
            "$orderby": "lastModifiedDateTime desc",
            "$select": "id,subject,bodyPreview,createdDateTime,lastModifiedDateTime",
        },
    ).json()
    return [
        {
            "id": m["id"],
            "title": m.get("subject") or "(untitled)",
            "preview": m.get("bodyPreview", ""),
            "created": m.get("createdDateTime"),
            "modified": m.get("lastModifiedDateTime"),
        }
        for m in data.get("value", [])
    ]


@mcp.tool()
def get_note(note_id: str) -> dict:
    """Get the full text of one synced note by id."""
    m = graph(
        "GET",
        f"/me/messages/{note_id}",
        params={"$select": "id,subject,body,createdDateTime,lastModifiedDateTime,hasAttachments"},
    ).json()
    body = m.get("body", {})
    text = body.get("content", "")
    if body.get("contentType") == "html":
        text = _strip_html(text)
    return {
        "id": m["id"],
        "title": m.get("subject") or "(untitled)",
        "text": text,
        "created": m.get("createdDateTime"),
        "modified": m.get("lastModifiedDateTime"),
        "has_attachments": m.get("hasAttachments", False),
    }


@mcp.tool()
def search_notes(query: str, folder_id: str, top: int = 25) -> list[dict]:
    """Full-text search within a notes folder (id from find_note_folders)."""
    data = graph(
        "GET",
        f"/me/mailFolders/{folder_id}/messages",
        params={
            "$search": f'"{query.replace(chr(34), "")}"',
            "$top": min(max(top, 1), 100),
            "$select": "id,subject,bodyPreview,lastModifiedDateTime",
        },
        headers={"ConsistencyLevel": "eventual"},
    ).json()
    return [
        {
            "id": m["id"],
            "title": m.get("subject") or "(untitled)",
            "preview": m.get("bodyPreview", ""),
            "modified": m.get("lastModifiedDateTime"),
        }
        for m in data.get("value", [])
    ]


# ---------------------------------------------------------------- tools: OneNote


@mcp.tool()
def list_onenote_sections() -> list[dict]:
    """List OneNote sections across all notebooks."""
    data = graph(
        "GET", "/me/onenote/sections", params={"$expand": "parentNotebook", "$top": 100}
    ).json()
    return [
        {
            "id": s["id"],
            "section": s.get("displayName"),
            "notebook": (s.get("parentNotebook") or {}).get("displayName"),
        }
        for s in data.get("value", [])
    ]


@mcp.tool()
def list_onenote_pages(section_id: str | None = None, query: str | None = None, top: int = 25) -> list[dict]:
    """List OneNote pages, optionally within one section or filtered by title text."""
    path = f"/me/onenote/sections/{section_id}/pages" if section_id else "/me/onenote/pages"
    params: dict[str, Any] = {
        "$top": min(max(top, 1), 100),
        "$orderby": "lastModifiedDateTime desc",
        "$select": "id,title,createdDateTime,lastModifiedDateTime",
    }
    if query:
        params["$filter"] = f"contains(tolower(title),'{query.lower().replace(chr(39), chr(39) * 2)}')"
    data = graph("GET", path, params=params).json()
    return [
        {
            "id": p["id"],
            "title": p.get("title"),
            "created": p.get("createdDateTime"),
            "modified": p.get("lastModifiedDateTime"),
        }
        for p in data.get("value", [])
    ]


@mcp.tool()
def get_onenote_page(page_id: str) -> dict:
    """Get the text content of a OneNote page."""
    content = graph("GET", f"/me/onenote/pages/{page_id}/content").text
    title = re.search(r"(?is)<title>(.*?)</title>", content)
    return {
        "id": page_id,
        "title": html.unescape(title.group(1)) if title else None,
        "text": _strip_html(content),
    }


@mcp.tool()
def create_onenote_page(section_id: str, title: str, body: str) -> dict:
    """Create a OneNote page (plain text body) in a section. Writes to OneNote,
    not back into Samsung Notes, which only syncs one way."""
    paras = "".join(f"<p>{html.escape(line)}</p>" for line in body.splitlines() or [""])
    doc = (
        "<!DOCTYPE html><html><head>"
        f"<title>{html.escape(title)}</title></head><body>{paras}</body></html>"
    )
    p = graph(
        "POST",
        f"/me/onenote/sections/{section_id}/pages",
        content=doc.encode(),
        headers={"Content-Type": "text/html"},
    ).json()
    return {"id": p.get("id"), "title": p.get("title"), "url": (p.get("links") or {}).get("oneNoteWebUrl", {}).get("href")}


def main() -> None:
    if len(sys.argv) > 1 and sys.argv[1] == "login":
        get_token(interactive=True)
        print("Signed in; token cached at", CACHE_PATH, file=sys.stderr)
        return
    mcp.run()


if __name__ == "__main__":
    main()
