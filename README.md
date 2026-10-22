# 📧 Mail for Outlook

Share an Obsidian note, or part of it, without leaving Obsidian: open it as an **Outlook email draft**, or **copy it as rich text** to paste into OneNote, Jira, Teams or any app that accepts formatted text. And the other way round: **import an email** copied from Outlook as a searchable note, with its people linked to your person notes.

Before anything leaves your vault, the note is cleaned up for its readers: private parts are removed, queries become their results, embedded notes are filled in and links into your vault become plain text.

🖥️ Desktop only. Built for Microsoft Outlook (see Limitations below).

Based on [obsidian-mail-plugin](https://github.com/droezel/obsidian-mail-plugin) by **Dries Verbruggen**, released under the Apache License 2.0. See License and credits below.

## ✨ Features

- 📤 **Send a note or a selection** as a formatted HTML email draft.
- 📋 **Copy a note or a selection as rich text**, ready to paste into OneNote, Jira, Confluence, Teams or Word.
- 🖼️ **Images** from your vault travel with the content: inline attachments in emails, embedded in copies.
- 🧩 **Embedded notes** (`![[Other note]]`, `![[Other note#Heading]]`) are replaced by their content.
- 🔗 **Links into your vault** are turned into plain text: they mean nothing outside your vault.
- 👥 **Recipients filled in automatically** from linked person notes and from email addresses written in the note.
- 🛑 **Safety against early sending**: an optional placeholder in the To field stops Outlook from sending until you remove it.
- 🔒 **Privacy filter**: parts of a note tagged `#private` never leave your vault.
- 📊 **Dataview and Tasks queries** are replaced by their results.
- ☑️ **Task checkboxes** are shown as ☐ and ☑, which every app can display.
- ⏳ **Progress at a glance**: one notice shows the export's progress and turns into ✅ when it's done.
- 📥 **Import an email** from Outlook as a note: header turned into properties, people linked to person notes, duplicates detected. Works with a Templater template, or without Templater.

## 📋 Why copy as rich text?

An email is only one way to share a note. Often the note needs to land somewhere else: the meeting minutes in a team OneNote notebook, a specification in a Jira ticket or a Confluence page, a summary in a Teams chat, a section of a Word document.

**Copy note as rich text** prepares the note exactly as for an email (privacy filter, embedded notes, query results, vault links flattened) and puts the result on the clipboard. You then paste it wherever you need it, with headings, lists, tables, links and formatting intact. No copy-and-clean-up by hand, and no risk of pasting a `#private` paragraph by mistake.

## 📦 Installation

### Manual

1. Download `mail-for-outlook_v<version>.zip` from the latest release and unzip it.
2. Copy the `mail-for-outlook` folder into your vault's `.obsidian/plugins/` folder (replace it if it already exists).
3. In Obsidian, go to **Settings → Community plugins**, then turn **Mail for Outlook** on (or off and on again after an update).

### BRAT

Add this repository `SpinEchoArcanist/obsidian-mail-for-outlook` in BRAT ("Add a beta plugin for testing") and pick a version.

## 🚀 Usage

| How | What it does |
|---|---|
| 📧 Ribbon icon **Send note as email** | Opens the whole active note as an email draft |
| ⌨️ Command **Mail for Outlook: Send note as email** | Opens the whole active note as an email draft |
| ⌨️ Command **Mail for Outlook: Send selection as email** | Opens only the selected text as an email draft (available when something is selected) |
| ⌨️ Command **Mail for Outlook: Copy note as rich text** | Copies the whole active note to the clipboard |
| ⌨️ Command **Mail for Outlook: Copy selection as rich text** | Copies only the selected text to the clipboard (available when something is selected) |
| ⌨️ Command **Mail for Outlook: Import email** | Creates a note from the email copied from Outlook, or a blank email note when there is none (see Importing an email) |

For an email, the subject is the note's name. The draft opens in your email client; nothing is sent until you press Send.

### ⏳ Progress notice

While an export runs, a notice stays on screen and fills up stage by stage:

`Copying ▰▰▰▱▱ Rendering…`

When it's done, the same notice changes to the result and closes by itself:

| Result | Notice | Stays |
|---|---|---|
| Done | ✅ Copied — ready to paste (or ✅ Email draft opened with 3 recipients) | 4 s |
| Done, with something left out | ✅ …, followed by ⚠️ Left out: 1 DataviewJS query (not supported) | 10 s |
| Failed | ❌ Copy failed: reason | 8 s |

The number of private areas removed is added to the result line. While an export is running, starting another one does nothing.

## 📋 Copying as rich text

The copy commands go through exactly the same steps as an email: the privacy filter, embedded notes, query expansion, vault links turned into text, checkboxes as ☐/☑ and the **Include frontmatter** setting all apply. Only the last step differs: instead of opening a draft, the result is put on the clipboard.

- 👥 **No recipients**: the recipient settings and the safety placeholder only concern emails.
- 🖼️ **Images** from your vault are written into the copied content itself, since a clipboard can't carry attachments. They paste into OneNote and Word. Web apps such as Jira or Confluence usually only accept images uploaded as attachments, so images may be dropped there. Images from the web stay as links to the web.
- 📝 **Plain-text fallback**: apps that only accept plain text receive the visible text, without formatting.
- 🎨 **Formatting** depends on the app you paste into: headings, lists, tables, links and bold or italic text are usually kept; colours, code styling and callouts may be simplified.

## 📥 Importing an email

Archive an email from Outlook as an Obsidian note, so its text is searchable and linked to the people involved. The email itself stays in Outlook: the note keeps its subject and sent time so you can find the original (attachments, images) there.

### Setup (once)

1. Copy `templates/Email Template.md` from the release zip into your Templater templates folder, and set its path in **Settings → Mail for Outlook → Import → Email template (Templater)**, for example `Templates/Email Template.md`.
2. Optional, for when Templater is not available: copy `templates/Email Template (fallback).md` too and set it in **Fallback template**. Without it, the same built-in content is used.
3. Give **Mail for Outlook: Import email** a hotkey in **Settings → Hotkeys**.

### Every time

1. In Outlook (classic, new or web): open the email → **Forward** → **Ctrl+A** → **Ctrl+C** → discard the draft. Forward is used because it makes Outlook write the full header (De / Envoyé / À / Cc / Objet, or From / Sent / To / Cc / Subject). Your signature in the draft doesn't matter: everything above the header is dropped.
2. In Obsidian, press your hotkey. Confirm the title (prefilled with the subject), and the people window if it opens.

The note is created in `2 Mail/Liste Emails/`, named `✉️ <title>`, with these properties: `Type`, `Creation date`, `Envoyé:`, `De:`, `A:`, `Cc:`, `Objet`, `Tags`, `Comment`, then a `# [[note name]]` heading and the email from its header down. Images are not imported. Folder, prefix, properties and layout are decided by the template: to change them, edit the template.

Run with no Outlook email on the clipboard, the same command creates a **blank email note** with the same properties.

### What the plugin reads

- 🔎 **The header**: the first From line followed, within the next 6 lines, by a Sent and a Subject line. Older headers lower in the thread are ignored. French and English labels, with or without a space before `:`.
- 📅 **The sent date**: French (`mercredi 29 juillet 2026 19:30`), English (`Wednesday, July 29, 2026 7:30 PM`) or numeric, day first (`29.07.2026 19:30`, `29/07/2026 19:30`). A date it can't read leaves `Envoyé:` empty.
- 👥 **People** in From, To and Cc, separated by `;`. A comma alone doesn't separate people, so `Martin, Paul <paul.martin@example.com>` stays one person.

### 👥 People and person notes

Each person becomes a link to their person note (`[[Paul Martin]]`), otherwise to their address (`[[paul.martin@example.com]]`), otherwise their name as plain text when Outlook showed no address. Person notes are the same as for recipients: notes whose `type` property contains **Person note type**, with addresses in **Email property**.

1. **By address**: the address is in a person note's Email property (written as `x`, `mailto:x` or `[x](mailto:x)`) → linked directly.
2. **By name**: the name matches a person note's name or one of its aliases, in any word order, ignoring accents, case, commas and codes in brackets. `Martin Paul (ABC)` matches the note `Paul Martin`.
3. **Closest names**: if no name matches exactly, up to 3 notes with a similar name are offered, to catch spelling variants (`Martinez` / `Martines`).

People found by name (2 or 3) are shown in **one window per import**: for each, choose the note (or **No note**) and whether to **Add this address to the note** (on by default for an exact name, off for a similar name; not shown when there is no address). **Confirm** applies your choices: the address is added to the note's Email property (a single value becomes a list; nothing else in the note changes) and a notice tells you how many person notes were updated, so the next import matches them by address. **Skip** or Escape changes nothing: those people get their address link. People with no candidate at all (external contacts) get their address link without asking.

### Already archived? Numbering

- If a note with the same name and the same sent time is already in the folder, you are asked: **⛔ Annuler** or **✅ Importer quand même** (the copy then gets a number).
- Otherwise, a name already taken gets the next free number without asking: a later reply in the same thread becomes `✉️ RE Budget - 2`, then `- 3`.

### With or without Templater

- **Templater enabled** and **Email template** set: the command creates a new note from that template (Templater's "Create new note from template"). The template can also be run from Templater directly, with "Create new note from template" or "Insert template" on a new, empty note; text already in the note is kept below the heading.
- **Templater not installed, disabled, or no Email template set**: a notice says *Templater indisponible — modèle de secours utilisé* and the plugin creates the note itself from the **fallback template** (or the built-in default), in **Destination folder**, named **File name prefix** + title. The result has the same properties.
- If Templater fails during the import, the error is shown and no fallback note is made, so you never get two notes.

**Fallback template tokens** (`{{token}}`, no code): `{{title}}` (final note name), `{{subject}}`, `{{sent}}`, `{{from}}`, `{{to}}`, `{{cc}}`, `{{body}}`, `{{today:FORMAT}}` (a moment.js format, default `YYYY-MM-DD`). Inside the properties, tokens are written as YAML values: text in double quotes, people as a list of links, so subjects with quotes or colons stay valid. In the body, text is written as is and people as comma-separated links. Unknown `{{…}}` and Templater `<% … %>` are left unchanged.

### 🧩 For template authors

Templates reach the plugin like this (same pattern as Dataview):

```js
const mail = app.plugins.plugins["mail-for-outlook"]?.api;   // undefined if the plugin is off
const email = mail ? await mail.readEmailFromClipboard() : null;
```

| Member | Description |
|---|---|
| `version` | `1` |
| `readEmailFromClipboard()` | `null` if the clipboard holds no Outlook header; otherwise `{ subject, sent, sentRaw, from, to, cc, header, body }`. `sent` is `"YYYY-MM-DD HH:mm"` or `null`; `from` / `to` / `cc` are lists of `{ name, address, link }`; `header` is the header block alone and `body` the email from the header down, as markdown. Opens the people window when needed. |
| `findArchivedName({ folder, baseName, sentProperty, sent })` | `{ name, duplicate }`: `name` is `baseName` or `baseName - N` with the lowest free N; `duplicate` is a note in `folder` named like that whose `sentProperty` equals `sent` (or `null`). |

## 👥 Recipients

Recipients are added to the **To** field of an email draft, without duplicates. Your own address (the "Default from address" setting) is never added.

**From linked person notes.** Every linked note whose `type` property is `Personne` (configurable) and whose `Email` property (configurable) holds a valid address becomes a recipient, with the person's name. Links count wherever they are: in the text, in embedded notes and in the note's properties (for example an `attendees` property).

**From addresses in the text.** Every email address written in the note or in its properties is added, including `mailto:` links. Addresses inside code, comments, web links and file names are ignored. This makes the plugin useful without person notes.

**When sending a selection**, the setting **Recipients when sending a selection** decides where recipients come from:
- *Selection only* (default): people and addresses in the selected text, plus the note's properties.
- *Whole note*: everyone in the note, as for a full export.

🛑 **Safety against early sending.** When on, a placeholder that is not an email address (default `REMOVE-BEFORE-SENDING`) is put first in the To field. Outlook can't resolve it and refuses to send until you delete it, so a half-written email can't go out by mistake.

## 🔒 Privacy filter

Mark the parts of a note that must never leave your vault with the tag `#private`. The filter is always on, for emails and copies alike.

1. **On a list item** (`-`, `*`, `+`, `1.`, task items included), `#private` hides only that item and what belongs to it: its sub-items and the lines indented below it. Its siblings, its parent and the heading above stay.
2. **Anywhere else**, `#private` starts a private area at the line that contains it.
3. If that line is the first line under a heading (blank lines don't count), the heading is private too. A `#private` written on the heading line itself works the same way.
4. The area ends at a line containing `#privateEnd` (that line is private too). Without `#privateEnd`, it ends at the next heading of the same or higher level, so subsections are private too. Before the first heading of a note, it ends at the first heading.
5. A forgotten `#privateEnd` never leaks anything: the area simply runs to the end of the section.
6. `#private` and `#privateEnd` never appear in the result. Tags are matched without regard to case; `#privately` or `#private/sub` don't count; tags inside code are ignored.
7. The filter also applies inside embedded notes, including a section embedded from inside another note's private area, and inside a selection.
8. Private content is removed before anything else happens: its text, images, queries, linked people and addresses never reach the result. The final notice tells you how many private areas were removed.

Example:

```markdown
## Décisions
Visible to everyone.

## Notes internes
#private
Never shared, nor are its subsections.

## Actions
Visible.
#private
Hidden from here…
#privateEnd
…and visible again from here.
- Call the supplier
- Negotiate the discount #private
    - Our limit is 15 %
- Send the report
```

Here only "Negotiate the discount" and its sub-item are hidden; the two other actions stay.

💬 Obsidian comments (`%% … %%`) are not shown in the result either.

## 📊 Queries

With **Expand queries** on (default), queries are replaced by their results:

| Query | Result |
|---|---|
| ` ```dataview ` blocks: `TABLE`, `LIST`, `TASK` | Converted through Dataview's own API; works inside callouts too |
| Inline Dataview: `` `= expression` `` | Replaced by its value |
| ` ```tasks ` blocks | Rendered by Tasks, then cleaned: buttons, query explanation, task count and backlinks are removed |

🚫 Not supported: Dataview `CALENDAR` queries, `dataviewjs` blocks and inline DataviewJS (`` `$= …` ``). A query that can't be expanded (plugin missing or disabled, error, timeout after 5 seconds for Tasks) is left out, and the final notice lists what was left out. Without Dataview installed, inline `` `= …` `` code is kept as written.

A query without results (an empty table, list or Tasks query) is left out entirely; the heading you wrote above it stays.

Links in query results become plain text. With **Expand queries** off, query blocks are kept as their source code.

## ⚙️ Settings

A short description at the top of the settings reminds what the plugin does. The settings are grouped by what they affect; settings that depend on a switch only appear when that switch is on.

| Setting | Default | Description |
|---|---|---|
| **📄 Content** (emails and copies) | | |
| Include frontmatter | off | Include the note's YAML frontmatter |
| Expand queries | on | Replace Dataview and Tasks queries with their results |
| **📧 Email** | | |
| Default from address | empty | Sender address; also never added as a recipient |
| Safety against early sending | off | Put a placeholder first in the To field |
| Safety placeholder | `REMOVE-BEFORE-SENDING` | Shown when safety is on. Spaces become dashes, `@` and address punctuation are removed |
| **👥 Recipients** (email drafts only) | | |
| From linked person notes | on | Add linked people's addresses |
| Person note type | `Personne` | Value of the `type` property that marks a person note. Also used by Import email |
| Email property | `Email` | Property holding their addresses (name matched without regard to case). Also used by Import email, which adds the addresses you confirm |
| From addresses in the text | on | Add addresses written in the note or its properties |
| Recipients when sending a selection | Selection only | Selection only, or Whole note |
| **📥 Import** | | |
| Email template (Templater) | empty | Path of the Templater template used by Import email |
| Fallback template | empty | Path of the `{{token}}` template used without Templater; empty = built-in default |
| Destination folder | `2 Mail/Liste Emails` | Fallback only (the Templater template sets its own) |
| File name prefix | `✉️ ` | Fallback only |
| Sent property | `Envoyé:` | Property holding the sent date, for the fallback's duplicate check |

## 🔧 How it works

1. The private areas are removed and embedded notes are inserted.
2. Recipients are collected (emails only).
3. Dataview queries are replaced by their results.
4. The note is rendered with Obsidian's own Markdown renderer, in an invisible area of the window so that Tasks can draw its results.
5. Vault links become text, checkboxes become ☐/☑, images become inline attachments (emails) or are written into the content (copies).
6. 📧 For an email, an `.eml` draft (with `X-Unsent: 1`) is written to the system's temporary folder, in a `mail-for-outlook` subfolder, and opened with your default email client. Drafts are deleted automatically after 5 minutes, or at the next start of Obsidian.
   📋 For a copy, the result is put on the clipboard as rich text, with plain text as a fallback; no file is written.

📥 For **Import email**, the clipboard is read only when you run the command; its HTML is converted with Obsidian's own converter after images are removed.

🛡️ The plugin makes no network requests. Exports never modify your notes; Import email creates the email note and, only when you confirm it, adds addresses to person notes.

## ⚠️ Limitations

- 📧 **Email clients**: the `X-Unsent: 1` header that opens the file as an editable draft is respected by Outlook; Thunderbird may open it read-only.
- 🖼️ **SVG images** are attached but Outlook does not display them inline.
- 📋 **Images in copies**: whether the app you paste into keeps them depends on that app. Large images make the copied content heavy.
- 🪟 **Copying needs Obsidian in focus**: if you switch to another window while a copy is being prepared (Tasks queries can take a few seconds), the copy fails with a ❌ notice; run the command again.
- 📊 **Tasks expansion** relies on the class names Tasks uses in its results; a future Tasks update could break it. The query is then left out with a notice; the result itself is not broken.
- 🔒 **Privacy and queries**: the privacy filter works on the exported note's own text. Tasks or rows that a query collects from private areas of *other* notes are not filtered.
- 🧩 **Embeds inside code blocks** (`![[…]]` written inside a fenced code block) are also expanded.
- 🖼️ **Images inside embedded notes** that use a path relative to that note may not be found.
- 📥 **Import email and Templater**: the plugin calls Templater's internal `create_new_note_from_template` function, which is not a documented API. It checks the function is there before each call; if a Templater update removes it, the fallback template is used instead.
- 📥 **Import email** reads only the text: attachments, images, `.msg` and `.eml` files are not imported. The clipboard can only be read while Obsidian has the focus.
- 📥 **Cancelling** (at the title or at the duplicate question) when the Import email command runs through Templater leaves the new, empty note Templater had already created; delete it if you don't need it. The fallback creates nothing when you cancel.

## 🛠️ Building from source

```bash
npm install
npm run build   # type-checks and writes main.js
npm run dev     # rebuilds main.js on every change
```

`main.js` is not kept in the repository. To publish a version, create a GitHub release whose tag is the version number (for example `1.3.1`) and attach `main.js` and `manifest.json` to it; BRAT installs from those release files.

## 📜 License and credits

Licensed under the [Apache License 2.0](LICENSE).

- © 2026 SpinEchoArcanist.
- Based on [obsidian-mail-plugin](https://github.com/droezel/obsidian-mail-plugin) © Dries Verbruggen, also licensed under the Apache License 2.0. The changes made to the original are listed in the changelog below and in the notice at the top of `main.ts`.

## 📝 Changelog

### 1.4.0
- New command **Import email**: turns an email copied from Outlook (Forward → Ctrl+A → Ctrl+C) into a note through a Templater template, or through a built-in `{{token}}` fallback template when Templater is not available. Without an email on the clipboard, it creates a blank email note.
- People in From, To and Cc linked to person notes by address, by name (any word order, aliases) or by closest name, with a review window that can add confirmed addresses to the person notes.
- Duplicate check on subject and sent time; numbered names (`- 2`, `- 3`) for other emails with the same subject.
- API for templates: `app.plugins.plugins["mail-for-outlook"].api` with `readEmailFromClipboard()` and `findArchivedName()`.
- New settings section **Import**. **Person note type** and **Email property** are now always shown, since Import email uses them too.
- Templates shipped in the release zip: `Email Template.md` (Templater) and `Email Template (fallback).md`.

### 1.3.1
- Renamed to **Mail for Outlook**: new plugin id `mail-for-outlook`, command names "Send note as email", "Send selection as email", "Copy note as rich text" and "Copy selection as rich text". Settings and hotkeys from the "Mail" versions don't carry over.
- Settings grouped into Content, Email and Recipients, with a short description at the top; dependent settings only appear when their switch is on.
- One progress notice per export, ending in ✅ (4 s), ⚠️ with what was left out (10 s) or ❌ (8 s).
- A second export can't start while one is running.

### 1.3.0 
- New commands "Copy note as rich text" and "Copy selected content as rich text": the same export put on the clipboard, to paste into OneNote or other apps. Images are written into the copied content; no recipients.

### 1.2.1 
- Queries without results are left out instead of showing an empty table or block.
- `#private` on a list item hides only that item and its sub-items.

### 1.2.0 
- Recipients from email addresses written in the note and its properties; person-note links in properties now count too.
- New settings: switches for each recipient source; recipients when sending a selection (selection only / whole note).
- Safety against early sending: optional placeholder in the To field.
- Privacy filter with `#private` and `#privateEnd`.
- Dataview (blocks and inline) and Tasks queries expanded; setting to turn it off.
- Task checkboxes shown as ☐ / ☑; tables get borders.

### 1.1.0 
- Images sent as inline attachments instead of base64 data.
- New command "Send selected content".
- Links into the vault turned into plain text; embedded notes expanded.
- Recipients from linked person notes.
- Exported drafts kept in a dedicated temp folder and deleted automatically.
- Accented subjects encoded correctly; frontmatter removal works with Windows line endings; code-quality fixes from Obsidian's linter.

### 1.0.1 and 1.0.0
- Original releases of obsidian-mail-plugin by Dries Verbruggen: export a note as an HTML email draft, sidebar button, image embedding, sender and frontmatter settings.
