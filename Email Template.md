<%*
// ========================================================
// EMAIL TEMPLATE — Mail for Outlook 1.4.0
// ========================================================
/*
Archives an email copied from Outlook as a note.

HOW TO USE
1. Outlook: open the email → Forward → Ctrl+A → Ctrl+C → discard the draft.
2. Obsidian: run "Mail for Outlook: Import email" (set this template in the plugin's
   settings, Import → Email template). Confirm the title, and the people window if it opens.
   Without an Outlook email on the clipboard, the same command makes a blank email note.
   The template can also be run from Templater ("Create new note from template", or
   "Insert template" on a new empty note).

WHO DOES WHAT
- Mail for Outlook reads the clipboard, finds the header, links people to person notes and
  finds the archive name (duplicate check and "- 2", "- 3" numbering).
- This template decides the note: folder, name prefix, title prompt, properties, layout.
  Change the format here; the plugin doesn't need to change.

Properties and body are written once Templater has finished (tp.hooks), so Templater's own
writing of the template output can't overwrite them, whichever way the template is run.
*/

// 🔧 CONFIGURATION
const DESTINATION = "2 Mail/Liste Emails";   // Folder of archived emails
const PREFIX = "✉️ ";                        // Put before the title in the note name
const SENT_PROPERTY = "Envoyé:";             // Property holding the sent date (duplicate check)
// Removed from titles: not allowed in file names, or breaking the "# [[…]]" heading link
const ILLEGAL = /[\\/:*?"<>|#@%;[\]^]/g;

// ========================================================
// STEP 1: READ THE EMAIL (plugin missing or disabled, or no email → blank note)
// ========================================================
const mail = app.plugins.plugins["mail-for-outlook"]?.api;
const email = mail ? await mail.readEmailFromClipboard() : null;
const file = tp.config.target_file;

// ========================================================
// STEP 2: TITLE
// ========================================================
const prefill = email ? email.subject.replace(ILLEGAL, "").trim() : "";
const answer = await tp.system.prompt("Titre de l'email (Objet) :", prefill);
if (answer === null) {   // Cancel / Escape: the note is left untouched
  new Notice("⛔ Import annulé");
  return;
}
const title = answer.replace(ILLEGAL, "").trim() || tp.file.title;
const baseName = PREFIX + title;

// ========================================================
// STEP 3: NAME — duplicate check and numbering
// ========================================================
let name;
if (mail) {
  const found = mail.findArchivedName({
    folder: DESTINATION,
    baseName,
    sentProperty: SENT_PROPERTY,
    sent: email ? email.sent : null
  });
  if (found.duplicate) {
    const choice = await tp.system.suggester(
      ["⛔ Annuler", "✅ Importer quand même"],
      ["cancel", "import"],
      false,
      `Cet e-mail est déjà archivé : ${found.duplicate.basename}`
    );
    if (choice !== "import") {
      new Notice("⛔ Import annulé");
      return;
    }
  }
  name = found.name;   // Used without asking: a later RE: gets "- 2", "- 3"…
} else {
  // Plugin missing: same rule, lowest free "- N" from 2
  name = baseName;
  let n = 2;
  while (app.vault.getAbstractFileByPath(`${DESTINATION}/${name}.md`)) {
    name = `${baseName} - ${n++}`;
  }
}

// ========================================================
// STEP 4: RENAME AND MOVE
// ========================================================
await tp.file.rename(name);
await tp.file.move(`${DESTINATION}/${name}`);

// ========================================================
// STEP 5: PROPERTIES AND BODY — once Templater has finished writing the note
// ========================================================
const links = (people) => people.map(p => p.link);
const properties = {
  "Type": ["e-mail"],
  "Creation date": tp.date.now("YYYY-MM-DD"),
  "Envoyé:": email && email.sent ? email.sent : "",
  "De:": email ? links(email.from) : [],
  "A:": email ? links(email.to) : [],
  "Cc:": email ? links(email.cc) : [],
  "Objet": email ? email.subject : "",
  "Tags": [],
  "Comment": ""
};

tp.hooks.on_all_templates_executed(async () => {
  try {
    // Frontmatter: these keys first, in this order; other properties already there are kept after them
    await app.fileManager.processFrontMatter(file, (frontmatter) => {
      const ours = Object.keys(properties).map(key => key.toLowerCase());
      const others = {};
      for (const key of Object.keys(frontmatter)) {
        if (!ours.includes(key.toLowerCase())) others[key] = frontmatter[key];
        delete frontmatter[key];
      }
      Object.assign(frontmatter, properties, others);
    });

    // Body: heading, then any text the note already had, then the email
    await app.vault.process(file, (data) => {
      const info = tp.obsidian.getFrontMatterInfo(data);
      const head = info.exists ? data.slice(0, info.contentStart) : "";
      const existing = data.slice(head.length).trim();
      const parts = [`# [[${name}]]`];
      if (existing) parts.push(existing);
      if (email && email.body) parts.push(email.body);
      return `${head}\n${parts.join("\n\n")}\n`;
    });

    new Notice(`✅ E-mail archivé\n📁 ${DESTINATION}\n✏️ ${name}`);
  } catch (error) {
    new Notice(`❌ Import : ${error.message}`, 8000);
    console.error("Email Template error:", error);
  }
});
-%>
