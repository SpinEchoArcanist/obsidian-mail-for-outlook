# 📧 Mail for Outlook

Share an Obsidian note, or part of it, without leaving Obsidian: open it as an **Outlook email draft**, or **copy it as rich text** to paste into OneNote, Jira, Teams or any app that accepts formatted text.

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
| Person note type | `Personne` | Shown when the option above is on. Value of the `type` property that marks a person note |
| Email property | `Email` | Shown when person notes are on. Property holding their address (name matched without regard to case) |
| From addresses in the text | on | Add addresses written in the note or its properties |
| Recipients when sending a selection | Selection only | Selection only, or Whole note |

## 🔧 How it works

1. The private areas are removed and embedded notes are inserted.
2. Recipients are collected (emails only).
3. Dataview queries are replaced by their results.
4. The note is rendered with Obsidian's own Markdown renderer, in an invisible area of the window so that Tasks can draw its results.
5. Vault links become text, checkboxes become ☐/☑, images become inline attachments (emails) or are written into the content (copies).
6. 📧 For an email, an `.eml` draft (with `X-Unsent: 1`) is written to the system's temporary folder, in a `mail-for-outlook` subfolder, and opened with your default email client. Drafts are deleted automatically after 5 minutes, or at the next start of Obsidian.
   📋 For a copy, the result is put on the clipboard as rich text, with plain text as a fallback; no file is written.

🛡️ The plugin makes no network requests and never modifies your notes.

## ⚠️ Limitations

- 📧 **Email clients**: the `X-Unsent: 1` header that opens the file as an editable draft is respected by Outlook; Thunderbird may open it read-only.
- 🖼️ **SVG images** are attached but Outlook does not display them inline.
- 📋 **Images in copies**: whether the app you paste into keeps them depends on that app. Large images make the copied content heavy.
- 🪟 **Copying needs Obsidian in focus**: if you switch to another window while a copy is being prepared (Tasks queries can take a few seconds), the copy fails with a ❌ notice; run the command again.
- 📊 **Tasks expansion** relies on the class names Tasks uses in its results; a future Tasks update could break it. The query is then left out with a notice; the result itself is not broken.
- 🔒 **Privacy and queries**: the privacy filter works on the exported note's own text. Tasks or rows that a query collects from private areas of *other* notes are not filtered.
- 🧩 **Embeds inside code blocks** (`![[…]]` written inside a fenced code block) are also expanded.
- 🖼️ **Images inside embedded notes** that use a path relative to that note may not be found.

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
