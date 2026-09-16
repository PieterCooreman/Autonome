# Autonome Agentic CMS

**Websites that build themselves.**

Autonome is a self-hosted, AI-powered website builder. You describe your website in plain
language (purpose, audience, style, content) and an autonomous AI agent researches the topic,
designs and writes the complete site for you — as a single, self-contained `index.html` built
on Bootstrap 5. It runs entirely on your own hardware against your own LLM. No cloud, no
subscriptions, no data leaving your network.

Autonome is **built for ASPPY**, the Classic ASP/VBScript runtime for Python. The whole
application lives in this `www` folder: the UI is a React single-page app (`index.html` +
`script.js` + `style.css`), and the backend is the REST JSON API in `api.asp` backed by a
SQLite database (`app.db`) — served by the ASPPY runtime.

- ASPPY project: https://github.com/PieterCooreman/ASPPY
- ASPPY on PyPI (installable): https://www.piwheels.org/project/asppy/

The ASPPY source folder is **not** required — install the runtime with `pip` (see
[Requirements](#requirements) below).

---

## Table of contents

- [Features](#features)
- [Requirements](#requirements)
- [Start the application](#start-the-application)
- [First login](#first-login)
- [The user interface](#the-user-interface)
  - [Landing page](#landing-page)
  - [Account](#account)
  - [Dashboard (My Projects)](#dashboard-my-projects)
  - [Project workspace (AI Builder)](#project-workspace-ai-builder)
- [How a generation works](#how-a-generation-works)
- [Uploading images](#uploading-images)
- [Version history & restores](#version-history--restores)
- [Downloading your site](#downloading-your-site)
- [Admin panel](#admin-panel)
  - [LLM configuration](#llm-configuration)
  - [Email settings](#email-settings)
  - [Users](#users)
- [Where files are stored](#where-files-are-stored)
- [The API at a glance](#the-api-at-a-glance)
- [Tips for best results](#tips-for-best-results)
- [Security notes](#security-notes)

---

## Features

- **Agentic AI generation** — the agent doesn't just complete text; it researches your topic
  (optional Wikipedia tools) and writes a complete, responsive website.
- **Your hardware, your model** — works with LM Studio, Ollama (OpenAI-compatible endpoints)
  or Anthropic's API. No cloud requirement.
- **Plain-language editing** — after the first build, keep asking for changes ("make the hero
  full-screen", "add a testimonials section", "switch to a dark theme") and the AI patches the
  page.
- **Bring your own photos** — upload images; the AI weaves them into the design. Images wider
  than 1920 px are automatically resized.
- **Version history** — a backup is created automatically before every generation. Preview any
  previous version full-screen and restore it with one click (restores are also backed up, so
  nothing is ever lost).
- **Single-file output** — clean, semantic, mobile-first HTML/CSS/JS on Bootstrap 5.3 with a
  cookie notice and privacy statement included, plus Open Graph tags injected automatically so
  shared links show a proper preview.
- **Download & publish anywhere** — every project downloads as a ready-to-host `autonome.zip`.
- **Multi-user** — every user gets isolated projects and files. Admins manage the AI
  configuration centrally and can change user roles.
- **Async generation** — long builds run in the background. You can safely leave the page (or
  close the browser); the generation continues and you can pick it up again later. Optionally
  get an email when it finishes.
- **Self-contained & private** — a single folder of files, no npm, no build step, no external
  tracking.

---

## Requirements

- **Python 3.9+** installed on the server (Windows, Linux or macOS).
- **ASPPY** installed via pip: `pip install asppy`. The ASPPY source folder is **not**
  included with, or needed by, this app — Autonome runs entirely on the pip-installed runtime.
  (See https://github.com/PieterCooreman/ASPPY and https://www.piwheels.org/project/asppy/.)
- An LLM endpoint to point at. Recommended: a local server such as **LM Studio** or **Ollama**
  (OpenAI-compatible), or an Anthropic API key. See [LLM configuration](#llm-configuration).
- (Optional) An SMTP account if you want password-reset and generation-finished emails.

---

## Start the application

Serve this `www` folder with the ASPPY runtime. From anywhere (the ASPPY source folder is not
required):

```bash
asppy 0.0.0.0 8080 www
```

Or, on Windows, double-click `start_www.bat` — it stops anything already listening on port
8080, starts the server and opens `http://localhost:8080` in your browser.

```text
Usage:  asppy [host] [port] [docroot]
        asppy 0.0.0.0 8080 www
```

> In a source checkout the equivalent is `python ASPPY/server.py 0.0.0.0 8080 www`, but after
> `pip install asppy` the `asppy` command alone is enough.

The database (`app.db`) is created automatically on the first request. No setup step needed.

---

## First login

1. Open `http://localhost:8080`.
2. Click **Sign in** (top right) and sign in with the default administrator account:

   | Field | Value |
   |-------|-------|
   | Username | `admin` |
   | Password | `admin123` |

   > **Please change this password right away** (Account → Change password). The default
   > account is created on first run and is widely known.

3. As an admin you'll land on the **Admin panel**. Before you can generate a website you must
   configure the LLM connection (see [LLM configuration](#llm-configuration)). Other users can
   also register their own accounts from the sign-in page.
4. Open the **Dashboard**, create a project and start building.

---

## The user interface

The app is a single-page application. All navigation happens through URL fragments (`#/...`).

### Landing page

The public homepage (`/`) explains what Autonome is and how it works. Click **Sign in** or
**Start building** to continue. There is no registration link here — use the login page.

### Login / register / forgot password

The sign-in page has three modes:

- **Sign in** — existing username + password.
- **Register free** — create a new account. Requires a username, a valid email address and a
  password of at least 6 characters. Usernames must be plain ASCII. Registering makes you a
  regular **user** (not an admin).
- **Forgot password?** — enter your email address; if it is registered (and SMTP is
  configured), a reset link is emailed to you. The link (valid for 1 hour) leads to a "Set a
  new password" page. The response is intentionally neutral so email addresses cannot be
  probed.

Admin logins are routed to the Admin panel automatically; everyone else goes to the Dashboard.

### Account

Your account page (`#/account`) lets you:

- **Update your email address** — enter the new address and confirm with your current password.
- **Change your password** — current password + new password (min. 6 characters) + confirm.

### Dashboard (My Projects)

The dashboard lists all projects you own, most recently updated first. For each project you
can:

- **Open workspace** — click the project card.
- **+ New Project** — create a project; it gets its own URL and a starter "onboarding" page
  instantly.
- **Rename** — change the display name.
- **Copy** — duplicate the project (website files, photos and prompt/chat history; backups are
  not copied).
- **Notes** — attach a free-text note to the project.
- **Delete** — remove the project permanently (all files, backups and history).

### Project workspace (AI Builder)

The workspace (`#/project/<id>`) is a split screen:

- **Left: live preview.** An iframe showing the current site. Use the browser-style bar to
  **Copy** the site's URL or **Refresh** the preview. Uploaded photos appear as a thumbnail
  strip above the preview; hover a thumbnail to delete it.
- **Right: the AI Builder chat.** This is where you "talk" to the agent:

  - Type your request in the textarea. **Enter** sends; **Shift+Enter** inserts a newline.
  - **Suggestions / Suggest improvement** (button above the prompt) opens a pick-list of
    starter prompts on a fresh project, or grouped improvement ideas (tone, theme, layout,
    colors, content structure) on an existing one. Clicking one appends it to your prompt.
  - **Email me when the generation is complete** — opt into a notification email for the next
    generation.

- **Header actions** (desktop; on small screens these move into a hamburger menu):

  - **Reset** — start all over: removes the website, prompt history and backups. Your
    uploaded photos are kept. Refused while a generation is running.
  - **History** — opens the version-history panel (see below).
  - **Open site** — opens the live site in a new tab.
  - **Download** — downloads the project as `autonome.zip`.
  - **Upload images** — pick one or more photos (jpg, jpeg, png, gif, webp, bmp, avif).

- **Chat bubbles:** your prompts and the agent's replies are kept in a persistent chat history
  (per project). Assistant messages show the model's **thinking** ("🧠 thinking"), duration,
  token usage and speed. You can copy any of your own prompts back to the clipboard.

---

## How a generation works

1. You submit a prompt (optionally after uploading photos).
2. A **job** is created. If the server supports it, a detached worker runs it in the
   background; otherwise it runs synchronously. Either way the UI polls for status every few
   seconds and shows a spinner ("AI is building your website").
3. The agent builds the complete prompt: your request, the project name, the current
   `index.html` (so it can improve, not rebuild), your image paths, the last 5 prompts as
   context, and (if enabled) access to Wikipedia search tools. It may call tools, reason, and
   then writes exactly **one** fenced ```html block.
4. The server validates that the answer is a complete HTML document, cleans it (replaces long
   dashes, makes image paths relative), strips/re-numbers its internal block-ID markers,
   injects a responsive "safety net" (viewport meta + overflow guards) and Open Graph tags,
   and **backs up the previous version** into `backups\<timestamp>\`.
5. The new `index.html` is written and the preview refreshes. You can keep refining in plain
   language or roll back via **History**.

If the model returns a broken or incomplete page, the system retries up to 3 times with
adaptive "repair" hints and otherwise leaves your site untouched and explains the error.
Truncated answers (`finish_reason = length`) are reported with a hint to raise `max_tokens`.

You can **cancel** a running generation (button next to the spinner). Any partial result is
discarded; nothing is written. If you navigate away mid-generation, the app re-attaches to the
still-running job when you come back.

---

## Uploading images

1. In the project workspace, click **Upload images** (top right) — multiple files at once.
2. Allowed formats: **jpg, jpeg, png, gif, webp, bmp, avif**.
3. Images are stored in the project's `img/` folder with a timestamped name, and are
   automatically resized to a maximum width of **1920 px**.
4. The AI sees the image list as part of every prompt and uses the photos in the design.
   Thumbnails appear above the preview. A photo can be deleted from the thumbnail strip.

> The AI is instructed to reference images with the exact relative paths it is given
> (`img/<filename>`), so generated sites work even after download/moving to any host.

---

## Version history & restores

- A backup is created **before every generation** (and before every restore).
- Open **History** in the workspace to see all backups, newest first, each with a file count.
- **Preview** opens the backup full-screen in a modal (images resolve to the live project).
- **Restore** copies the chosen backup over the current site. The current version is backed up
  first, so a restore is always undoable.
- **Reset** (in the header) removes the site, history and backups but keeps your photos.

---

## Downloading your site

Click **Download** in the workspace. The server stages your `index.html` (plus any legacy
`style.css`/`script.js`/`favicon.ico` and all images in `img/`), zips them as `autonome.zip`
and serves it. No backups or temp files are included. The zip is ready to drop onto any static
host.

---

## Admin panel

Only administrators see the **Admin** tab in the header. The panel has three tabs.

### LLM configuration

| Setting | Description |
|---------|-------------|
| **LLM Endpoint URL** | Base URL of your LM Studio / Ollama server, e.g. `http://localhost:1234`. OpenAI-compatible. |
| **API Key** | Needed for LM Studio remote access (and any OpenAI/Anthropic key). |
| **Model Name** | Click **↻ Models** to fetch the model list from the endpoint (OpenAI `/v1/models`; Anthropic returns a curated list). You can also type a model manually with the **✎ Edit** toggle. |
| **Wikipedia tools** | On/off — lets the AI research topics via `search_wikipedia` / `get_wikipedia_page` while building. |
| **System Prompt** | Optional extra instructions prepended to every generation. |
| **Extra LLM Parameters** | Optional JSON, e.g. `"temperature": 0.7, "max_tokens": 8192`. |
| **Internal URL** | The address the app uses to reach *itself* for the async worker (bypasses a reverse proxy). Must match the port ASPPY listens on, e.g. `http://127.0.0.1:8058` behind IIS ARR. |

### Email settings

- **SMTP server** — host, port (`25`, `465` for SSL, `587` for STARTTLS), login, password,
  SSL/TLS toggle and **From** address.
- **Public URL of this app** — used to build links inside emails (e.g. the password reset
  link). Default `http://localhost:8080`.
- **Password reset email** — subject and body template; placeholders `{username}` and `{link}`.
- **Generation finished email** — subject and body template; placeholders `{username}`,
  `{project}`, `{status}` and `{link}`.

Emails are only sent for the password-reset flow and when a user ticks "Email me when the
generation is complete".

### Users

A table of every account (username, email, role, created date) where an admin can promote users
to **admin** or demote them back to **user**.

---

## Where files are stored

Relative to the `www` root:

```
www/
├── app.db            SQLite database (users, projects, config, jobs, chat, ...)
├── index.html        The Autonome SPA itself (login, dashboard, admin, workspace)
├── script.js         React UI code
├── style.css         UI styling
├── api.asp           REST JSON API (all actions)
├── favicon.ico       Copied into every new project folder
├── robots.txt        Copied into every new project folder
└── sites/
    └── <random-16-char-folder>/   One folder per project (name is random — usernames are
        │                            never used in URLs or paths)
        ├── index.html             The generated website
        ├── img/                   Uploaded photos (resized to max 1920 px)
        └── backups/
            └── <timestamp>/       A copy of the site before each generation/restore
```

The project folder name is a random 16-character string generated at creation time — it is
never derived from the project name or username, so URLs stay stable and do not leak identities.

---

## The API at a glance

The SPA talks to `api.asp?action=...` over JSON (`Content-Type: application/x-www-form-urlencoded`
with a `json=` field, or `multipart/form-data` for uploads). Main actions:

| Action | Purpose |
|--------|---------|
| `setup` | Initialise/seed the database (auto-called). |
| `login`, `logout`, `session`, `register` | Authentication. |
| `change_password`, `update_email`, `forgot_password`, `reset_password` | Account management. |
| `get_config`, `save_config`, `list_models` | Admin LLM/email configuration. |
| `get_projects`, `create_project`, `rename_project`, `copy_project`, `delete_project`, `reset_project`, `update_notes`, `get_notes` | Project management. |
| `generate_website`, `generate_worker`, `generate_status`, `cancel_generation`, `active_job`, `keepalive` | AI generation jobs. |
| `upload_image`, `list_images`, `delete_image` | Image handling. |
| `list_backups`, `restore_backup`, `preview_backup`, `download_project`, `get_project_files` | History and exports. |
| `get_chat_history` | Per-project chat log. |
| `admin_get_users`, `admin_update_user` | User administration. |
| `version`, `ping` | Health/version info. |

Sessions are cookie-based; the SPA pings `keepalive` every 5 minutes while logged in.

---

## Tips for best results

- **Be specific.** Mention goal, audience, colors, style, tone and the content you want.
- **Upload your photos first** so the agent can design around them.
- **Iterate.** Don't expect perfection in one shot — ask for refinements in plain language
  ("make the hero full-screen", "add a call-to-action at the end of every section").
- **Use suggestions** to discover useful improvement patterns.
- **Big changes**: ask for a full redesign explicitly; small changes are handled by patching
  individual blocks of the page, leaving the rest untouched.
- **Write in your own language.** The generated site is written in the same language as your
  request.
- **Raise `max_tokens`** if you see "cut off by the token limit" errors.
- **Broken images?** Every generated `<img>` carries an `onerror` fallback to a placeholder,
  so the page still looks reasonable.

---

## Security notes

- Change the default **admin / admin123** credentials immediately after first login.
- Passwords are stored as **bcrypt** hashes (ASPPY crypto). Reset tokens are single-use and
  expire after 1 hour.
- The API is session-authenticated; project operations verify ownership (`user_id`), backups
  and filenames are checked against path traversal, and project folders use random names so
  usernames never appear in URLs.
- The app makes outbound requests only to the LLM endpoint you configure and (optionally) to
  the public Wikipedia API.
- Only `index.html` is written by the AI. `ValidateFileExtension` refuses any other filename in
  AI output, and every generated document must pass a structural completeness check before it
  is saved — a broken AI response never overwrites your site.