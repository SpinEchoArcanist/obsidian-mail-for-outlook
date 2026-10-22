/*
 * Mail for Outlook — send Obsidian notes as Outlook email drafts, or copy them as rich text.
 * © 2026 SpinEchoArcanist, licensed under the Apache License 2.0.
 * Based on obsidian-mail-plugin © Dries Verbruggen (https://github.com/droezel/obsidian-mail-plugin),
 * licensed under the Apache License 2.0. Versions up to 1.3.0 were named "Mail".
 * Modified in 2026 for v1.1.0: images sent as inline attachments, "Send selected content"
 * command, internal links removed, recipients from linked person notes, embedded notes
 * expanded, temp-file cleanup and code-quality fixes.
 * Modified in 2026 for v1.2.0: recipients from addresses in the text and properties,
 * recipient scope for selections, safety placeholder against early sending, #private /
 * #privateEnd privacy filter, Dataview and Tasks query expansion, task checkboxes as symbols.
 * Modified in 2026 for v1.2.1: empty query results are left out; #private on a list item hides
 * only that item and what is indented below it.
 * Modified in 2026 for v1.3.0: "Copy note as rich text" and "Copy selected content as rich text"
 * commands put the same export on the clipboard, to paste into OneNote or other apps.
 * Modified in 2026 for v1.3.1: renamed to Mail for Outlook (plugin id, commands, author);
 * settings grouped into sections; one progress notice per export; a second export can't start
 * while one is running.
 * Modified in 2026 for v1.4.0: "Import email" command that turns an email copied from Outlook into a
 * note through a Templater template, with a built-in token template when Templater is missing; people
 * matched to person notes (by address, name or close name, with a review window); API for templates.
 */
import { App, Plugin, PluginSettingTab, Setting, Notice, MarkdownRenderer, TFile, TFolder, Component, Modal, ToggleComponent, normalizePath, FileSystemAdapter, Editor, MarkdownView, MarkdownFileInfo, getFrontMatterInfo, getLinkpath, resolveSubpath, htmlToMarkdown, parseFrontMatterAliases, moment } from 'obsidian';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';

// eslint-disable-next-line @typescript-eslint/no-require-imports -- Electron is provided by Obsidian's desktop runtime, not an npm dependency of this plugin
const { shell } = require('electron') as { shell: { openPath(path: string): Promise<string> } };

interface MailForOutlookSettings {
	defaultFrom: string;
	includeFrontmatter: boolean;
	expandQueries: boolean;
	recipientsFromPersonNotes: boolean;
	personType: string;
	emailProperty: string;
	recipientsFromText: boolean;
	selectionRecipientScope: 'selection' | 'note';
	safetyPlaceholderEnabled: boolean;
	safetyPlaceholder: string;
	// Import (1.4.0)
	emailTemplate: string;
	fallbackTemplate: string;
	importFolder: string;
	importPrefix: string;
	sentProperty: string;
}

// ---------------------------------------------------------------------------------------------
// Import (1.4.0): API contract shared with the Templater template
// ---------------------------------------------------------------------------------------------

interface EmailPerson {
	// Display name as shown by Outlook ("" if only an address)
	name: string;
	// null when Outlook showed only a name
	address: string | null;
	// "[[Note basename]]" if matched; else "[[address]]"; else the plain name (no brackets)
	link: string;
}

interface ParsedEmail {
	// As in the header, prefixes (RE:, TR:, FW:) kept
	subject: string;
	// "YYYY-MM-DD HH:mm", null if the date can't be parsed
	sent: string | null;
	// The date text as found
	sentRaw: string;
	// [] when the field is absent
	from: EmailPerson[];
	to: EmailPerson[];
	cc: EmailPerson[];
	// The header block alone, as markdown
	header: string;
	// Markdown from the header block (included) to the end
	body: string;
}

interface FindArchivedNameOptions {
	folder: string;
	baseName: string;
	sentProperty: string;
	sent: string | null;
}

interface ArchivedName {
	name: string;
	duplicate: TFile | null;
}

interface MailForOutlookApi {
	version: number;
	readEmailFromClipboard(): Promise<ParsedEmail | null>;
	findArchivedName(options: FindArchivedNameOptions): ArchivedName;
}

// A person as read from the header, before matching to person notes.
interface RawPerson {
	name: string;
	address: string | null;
}

// The header and body found on the clipboard, before matching.
interface ParsedHeader {
	subject: string;
	sent: string | null;
	sentRaw: string;
	from: RawPerson[];
	to: RawPerson[];
	cc: RawPerson[];
	header: string;
	body: string;
}

type HeaderField = 'from' | 'sent' | 'to' | 'cc' | 'subject';

// A person note, indexed once per import from the metadata cache.
interface PersonNote {
	file: TFile;
	// Lowercase addresses from the email property
	addresses: Set<string>;
	// Normalized words of the file name and of each alias
	names: string[][];
}

// One person to confirm in the review window.
interface ReviewRow {
	key: string;
	person: RawPerson;
	candidates: TFile[];
	// true for an exact name match, false for closest names
	exact: boolean;
}

interface ReviewChoice {
	file: TFile | null;
	addAddress: boolean;
}

// Templater's internal object (not a documented API): checked at every call.
interface TemplaterInternals {
	create_new_note_from_template?: unknown;
}

type CreateNoteFromTemplate = (template: TFile, folder?: string, filename?: string, openNewNote?: boolean) => Promise<unknown>;

// Where an export goes: an email draft, or the clipboard as rich text.
type ExportTarget = 'email' | 'clipboard';

interface ImageAttachment {
	cid: string;
	filename: string;
	mimeType: string;
	data: ArrayBuffer;
}

interface Recipient {
	// Empty for addresses found in the text that don't belong to a linked person note
	name: string;
	address: string;
}

interface LinkRef {
	link: string;
	sourcePath: string;
}

// Markdown ready to render, with what was gathered on the way.
interface PreparedContent {
	markdown: string;
	links: LinkRef[];
	privateAreas: number;
}

// Result of scanning a text for #private / #privateEnd. area[i] is 0 for a public line,
// otherwise the number of the private area the line belongs to.
interface PrivacyAnalysis {
	lines: string[];
	area: number[];
}

interface Fence {
	start: number;
	end: number;
	marker: string;
	lang: string;
	quote: string;
	quoteDepth: number;
}

// Minimal view of Dataview's public API (exposed by Dataview as window.DataviewAPI).
interface DataviewResult<T> {
	successful: boolean;
	value?: T;
	error?: string;
}

interface DataviewApi {
	settings: { inlineQueryPrefix: string; inlineJsQueryPrefix: string };
	queryMarkdown(source: string, originFile?: string): Promise<DataviewResult<string>>;
	evaluateInline(expression: string, origin: string): DataviewResult<unknown>;
	value: { toString(value: unknown, settings?: unknown): string };
}

declare global {
	interface Window {
		DataviewAPI?: DataviewApi;
	}
}

const DEFAULT_SETTINGS: MailForOutlookSettings = {
	defaultFrom: '',
	includeFrontmatter: false,
	expandQueries: true,
	recipientsFromPersonNotes: true,
	personType: 'Personne',
	emailProperty: 'Email',
	recipientsFromText: true,
	selectionRecipientScope: 'selection',
	safetyPlaceholderEnabled: false,
	safetyPlaceholder: 'REMOVE-BEFORE-SENDING',
	emailTemplate: '',
	fallbackTemplate: '',
	importFolder: '2 Mail/Liste Emails',
	importPrefix: '✉️ ',
	sentProperty: 'Envoyé:'
}

// Exported drafts live in their own temp subfolder and are deleted once this old.
const TEMP_FOLDER = path.join(os.tmpdir(), 'mail-for-outlook');
const TEMP_FILE_MAX_AGE_MS = 5 * 60 * 1000;
// Maximum nesting depth when expanding embedded notes (also guards against embed cycles).
const MAX_EMBED_DEPTH = 3;
// Tasks draws its results asynchronously: poll until they are complete, up to the timeout.
const TASKS_TIMEOUT_MS = 5000;
const TASKS_POLL_MS = 100;
const TASKS_STABLE_POLLS = 3;
const EMAIL_PATTERN = /^[^\s@<>",;]+@[^\s@<>",;]+\.[^\s@<>",;]+$/;
const ASCII_PATTERN = /^[\x20-\x7E]*$/;
const ADDRESS_IN_TEXT = /[A-Za-z0-9._%+'-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+/g;
const HEADING = /^ {0,3}(#{1,6})(?:[ \t]|$)/;
const LIST_ITEM = /^[ \t]*(?:[-*+]|\d{1,9}[.)])(?:[ \t]|$)/;
const INLINE_CODE = /(`+)(?:[^`]|[^`][\s\S]*?[^`])\1(?!`)/g;
const COMMENTS = /%%[\s\S]*?%%|<!--[\s\S]*?-->/g;
const TAG_PRIVATE = /(^|\s)#private(?![\p{L}\p{N}_/-])/iu;
const TAG_PRIVATE_END = /(^|\s)#privateend(?![\p{L}\p{N}_/-])/iu;
const TAG_PRIVATE_END_ALL = /(^|\s)#privateend(?![\p{L}\p{N}_/-])/giu;
const QUERY_LANGUAGES = new Set(['dataview', 'dataviewjs', 'tasks']);
// Progress notice: number of stages shown in its bar, and how long the final message stays.
const EXPORT_STAGES = 5;
const NOTICE_DONE_MS = 4000;
const NOTICE_WARNING_MS = 10000;
const NOTICE_ERROR_MS = 8000;

// Import (1.4.0)
const TEMPLATER_ID = 'templater-obsidian';
// The header block: a From line followed, within this many non-empty lines, by Sent and Subject
const HEADER_WINDOW = 6;
// Closest-name candidates: minimum similarity (0 to 1) and how many are offered
const NAME_MATCH_THRESHOLD = 0.75;
const MAX_NAME_CANDIDATES = 3;
// Above this many words, words are paired greedily instead of trying every pairing
const MAX_EXHAUSTIVE_WORDS = 6;
// Removed from titles: not allowed in file names, or breaking the "# [[…]]" heading link
const TITLE_ILLEGAL = /[\\/:*?"<>|#@%;[\]^]/g;
// Name for a note whose title was left empty
const UNTITLED = 'Untitled';
const HEADER_LABELS: Record<string, HeaderField> = {
	de: 'from', from: 'from',
	envoye: 'sent', sent: 'sent', date: 'sent',
	a: 'to', to: 'to',
	cc: 'cc',
	objet: 'subject', subject: 'subject'
};
// A label at the start of a header line ("Envoyé :", "Importance:"); any space before ":" is allowed
const LABEL_LINE = /^(\p{L}[\p{L} -]{0,30}?)\s*:\s*(.*)$/u;
const MONTHS: Record<string, number> = {
	janvier: 1, fevrier: 2, mars: 3, avril: 4, mai: 5, juin: 6,
	juillet: 7, aout: 8, septembre: 9, octobre: 10, novembre: 11, decembre: 12,
	january: 1, february: 2, march: 3, april: 4, may: 5, june: 6,
	july: 7, august: 8, september: 9, october: 10, november: 11, december: 12,
	jan: 1, feb: 2, mar: 3, apr: 4, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12
};
// Built-in fallback template, used when no fallback template is set
const DEFAULT_FALLBACK_TEMPLATE = [
	'---',
	'Type:',
	'  - e-mail',
	'Creation date: {{today:YYYY-MM-DD}}',
	'"Envoyé:": {{sent}}',
	'"De:": {{from}}',
	'"A:": {{to}}',
	'"Cc:": {{cc}}',
	'Objet: {{subject}}',
	'Tags: []',
	'Comment: ""',
	'---',
	'',
	'# [[{{title}}]]',
	'',
	'{{body}}',
	''
].join('\n');

export default class MailForOutlookPlugin extends Plugin {
	settings: MailForOutlookSettings;
	// True while an export runs, so a second one can't start
	private exporting = false;
	// Reached from templates as app.plugins.plugins["mail-for-outlook"]?.api
	api: MailForOutlookApi | undefined;

	async onload() {
		await this.loadSettings();

		this.api = {
			version: 1,
			readEmailFromClipboard: () => this.readEmailFromClipboard(),
			findArchivedName: (options: FindArchivedNameOptions) => this.findArchivedName(options)
		};

		// Remove drafts left over from a previous session.
		void this.cleanupTempFiles();

		// Add ribbon icon (sidebar button)
		this.addRibbonIcon('mail', 'Send note as email', async () => {
			await this.exportActiveNote();
		});

		this.addCommand({
			id: 'send-note',
			name: 'Send note as email',
			callback: async () => {
				await this.exportActiveNote();
			}
		});

		this.addCommand({
			id: 'send-selection',
			name: 'Send selection as email',
			editorCheckCallback: (checking: boolean, editor: Editor, ctx: MarkdownView | MarkdownFileInfo) => {
				const selection = editor.getSelection();
				if (!selection || !ctx.file) return false;
				if (!checking) {
					void this.exportNote(ctx.file, editor.getValue(), selection);
				}
				return true;
			}
		});

		this.addCommand({
			id: 'copy-note',
			name: 'Copy note as rich text',
			callback: async () => {
				await this.exportActiveNote('clipboard');
			}
		});

		this.addCommand({
			id: 'copy-selection',
			name: 'Copy selection as rich text',
			editorCheckCallback: (checking: boolean, editor: Editor, ctx: MarkdownView | MarkdownFileInfo) => {
				const selection = editor.getSelection();
				if (!selection || !ctx.file) return false;
				if (!checking) {
					void this.exportNote(ctx.file, editor.getValue(), selection, 'clipboard');
				}
				return true;
			}
		});

		this.addCommand({
			id: 'import-email',
			name: 'Import email',
			callback: async () => {
				await this.importEmail();
			}
		});

		this.addSettingTab(new MailForOutlookSettingTab(this.app, this));
	}

	onunload() {
		this.api = undefined;
	}

	async exportActiveNote(target: ExportTarget = 'email') {
		const activeFile = this.app.workspace.getActiveFile();
		if (!activeFile) {
			new Notice('No active file');
			return;
		}
		await this.exportNote(activeFile, await this.app.vault.read(activeFile), undefined, target);
	}

	// Exports the whole note, or only `selection` when given; `noteContent` is always the whole note.
	// The result opens as an email draft, or is copied to the clipboard as rich text.
	async exportNote(file: TFile, noteContent: string, selection?: string, target: ExportTarget = 'email') {
		// One export at a time: the running one keeps its progress notice on screen
		if (this.exporting) return;
		this.exporting = true;
		const toEmail = target === 'email';
		const progress = new ProgressNotice(toEmail ? 'Preparing email' : 'Copying');
		const component = new Component();
		const staging = activeDocument.body.createDiv({ cls: 'mail-for-outlook-staging' });
		try {
			progress.step(1, 'Filtering private parts…');

			// Properties are never part of the privacy filter; they are shown only if the setting asks for it
			let content = selection ?? noteContent;
			const frontmatter = getFrontMatterInfo(content);
			const frontmatterText = frontmatter.exists ? content.slice(0, frontmatter.contentStart) : '';
			content = content.slice(frontmatterText.length);

			// Remove private areas and replace embedded notes with their content
			const prepared = await this.prepareContent(content, file.path);

			// Recipients come from what is sent, or from the whole note when the settings ask for it;
			// a copy to the clipboard has none
			const recipientSource = toEmail && selection !== undefined && this.settings.selectionRecipientScope === 'note'
				? await this.prepareContent(this.removeFrontmatter(noteContent), file.path)
				: prepared;
			const recipients = toEmail ? this.collectRecipients(file, recipientSource) : [];

			// Replace queries with their results
			progress.step(2, 'Expanding queries…');
			const failures: string[] = [];
			let markdown = await this.expandQueries(prepared.markdown, file.path, failures);
			if (this.settings.includeFrontmatter) {
				markdown = frontmatterText + markdown;
			}

			// Convert to HTML. The container is attached (off-screen) because some plugins,
			// such as Tasks, only draw their results in a visible element.
			progress.step(3, 'Rendering…');
			component.load();
			await MarkdownRenderer.render(this.app, markdown, staging, file.path, component);
			if (this.settings.expandQueries) {
				await this.settleTasksQueries(staging, failures);
			}
			this.convertCheckboxes(staging);

			// Links into the vault are meaningless to recipients: keep their text only
			this.removeInternalLinks(staging);

			// Process images
			progress.step(4, 'Adding images…');
			const attachments = await this.attachImages(staging, file.path);
			if (!toEmail) {
				this.embedImages(staging, attachments);
			}

			// Create styled HTML
			const serializer = new XMLSerializer();
			const html = Array.from(staging.childNodes).map(node => serializer.serializeToString(node)).join('');
			const styledHtml = this.wrapWithStyles(html);

			progress.step(5, toEmail ? 'Opening the draft…' : 'Copying to the clipboard…');
			let message: string;
			if (toEmail) {
				// Create EML
				const emlPath = await this.createEML(
					file.basename,
					styledHtml,
					attachments,
					recipients
				);

				// Open in email client
				await shell.openPath(emlPath);
				this.scheduleTempCleanup();

				message = recipients.length > 0
					? `✅ Email draft opened with ${recipients.length} recipient${recipients.length === 1 ? '' : 's'}`
					: '✅ Email draft opened';
			} else {
				// Rich text for apps that accept it, with the visible text as a fallback for plain-text fields
				await navigator.clipboard.write([new ClipboardItem({
					'text/html': new Blob([styledHtml], { type: 'text/html' }),
					'text/plain': new Blob([staging.innerText], { type: 'text/plain' })
				})]);
				message = '✅ Copied — ready to paste';
			}
			if (prepared.privateAreas > 0) {
				message += `; ${prepared.privateAreas} private area${prepared.privateAreas === 1 ? '' : 's'} removed`;
			}
			// What couldn't be exported goes on a second line, and the message stays longer
			if (failures.length > 0) {
				progress.finish([message, `⚠️ Left out: ${this.summarizeFailures(failures)}`], NOTICE_WARNING_MS);
			} else {
				progress.finish([message], NOTICE_DONE_MS);
			}
		} catch (error) {
			progress.finish([`❌ ${toEmail ? 'Email' : 'Copy'} failed: ${error instanceof Error ? error.message : String(error)}`], NOTICE_ERROR_MS);
			console.error('Mail for Outlook export error:', error);
		} finally {
			component.unload();
			staging.remove();
			this.exporting = false;
		}
	}

	private async prepareContent(markdown: string, sourcePath: string): Promise<PreparedContent> {
		const prepared: PreparedContent = { markdown: '', links: [], privateAreas: 0 };
		const filtered = keepPublic(analyzePrivacy(markdown));
		prepared.privateAreas += filtered.areas;
		prepared.markdown = await this.prepareLevel(filtered.text, sourcePath, 1, new Set([sourcePath]), prepared);
		return prepared;
	}

	// `markdown` has already been through the privacy filter.
	private async prepareLevel(markdown: string, sourcePath: string, depth: number, visited: Set<string>, prepared: PreparedContent): Promise<string> {
		for (const link of extractLinks(markdown)) {
			prepared.links.push({ link, sourcePath });
		}
		return this.expandEmbeddedNotes(markdown, sourcePath, depth, visited, prepared);
	}

	private async expandEmbeddedNotes(markdown: string, sourcePath: string, depth: number, visited: Set<string>, prepared: PreparedContent): Promise<string> {
		const embedRegex = /!\[\[([^\]]+?)\]\]/g;
		const matches: RegExpExecArray[] = [];
		let found: RegExpExecArray | null;
		while ((found = embedRegex.exec(markdown)) !== null) {
			matches.push(found);
		}
		if (matches.length === 0) return markdown;

		let result = '';
		let lastIndex = 0;
		for (const match of matches) {
			const index = match.index;
			result += markdown.slice(lastIndex, index);
			lastIndex = index + match[0].length;

			const linktext = match[1].split('|')[0];
			const hashIndex = linktext.indexOf('#');
			const subpath = hashIndex >= 0 ? linktext.slice(hashIndex) : '';
			const target = this.app.metadataCache.getFirstLinkpathDest(getLinkpath(linktext), sourcePath);

			// Images and other attachments stay as embeds; unresolved or too-deep notes are dropped
			if (!target || target.extension !== 'md') {
				result += match[0];
				continue;
			}
			if (depth > MAX_EMBED_DEPTH || visited.has(target.path)) {
				continue;
			}

			// The privacy filter runs on the whole embedded note first, so a section embedded
			// from inside a private area stays private
			const raw = await this.app.vault.cachedRead(target);
			const info = getFrontMatterInfo(raw);
			const bodyOffset = info.exists ? info.contentStart : 0;
			const body = raw.slice(bodyOffset);
			const analysis = analyzePrivacy(body);
			let from = 0;
			let to = analysis.lines.length;
			const cache = this.app.metadataCache.getFileCache(target);
			const section = subpath && cache ? resolveSubpath(cache, subpath) : null;
			if (section) {
				from = lineIndexAt(body, Math.max(0, section.start.offset - bodyOffset));
				if (section.end) {
					const endOffset = Math.max(0, section.end.offset - bodyOffset);
					to = lineIndexAt(body, endOffset) + (endOffset === 0 || body[endOffset - 1] === '\n' ? 0 : 1);
				}
			}
			const filtered = keepPublic(analysis, from, to);
			prepared.privateAreas += filtered.areas;

			const nestedVisited = new Set(visited);
			nestedVisited.add(target.path);
			const embedded = await this.prepareLevel(filtered.text, target.path, depth + 1, nestedVisited, prepared);
			result += `\n\n${embedded.trim()}\n\n`;
		}
		result += markdown.slice(lastIndex);
		return result;
	}

	// Replaces Dataview queries with their results; Tasks queries are left for the renderer.
	private async expandQueries(markdown: string, sourcePath: string, failures: string[]): Promise<string> {
		const lines = splitLines(markdown);
		const fences = findFences(lines);
		const fenceAt = new Map(fences.map(fence => [fence.start, fence]));
		const api = this.settings.expandQueries ? window.DataviewAPI : undefined;
		const out: string[] = [];

		for (let i = 0; i < lines.length; i++) {
			const fence = fenceAt.get(i);
			if (!fence) {
				out.push(api ? this.expandInlineQueries(lines[i], api, sourcePath, failures) : lines[i]);
				continue;
			}
			const block = lines.slice(fence.start, fence.end + 1);
			i = fence.end;

			if (!QUERY_LANGUAGES.has(fence.lang)) {
				out.push(...block);
				continue;
			}
			if (!this.settings.expandQueries) {
				// Not expanded: show the query itself as plain code
				out.push(block[0].replace(/(`{3,}|~{3,})[ \t]*[^\s`]+/, '$1'), ...block.slice(1));
				continue;
			}
			if (fence.lang === 'tasks') {
				out.push(...block);
				continue;
			}
			if (fence.lang === 'dataviewjs') {
				failures.push('DataviewJS query (not supported)');
				continue;
			}

			// Dataview query
			if (!api) {
				failures.push('Dataview query (Dataview not available)');
				continue;
			}
			const closed = fence.end > fence.start && isFenceClose(lines[fence.end], fence);
			const source = block.slice(1, closed ? -1 : undefined)
				.map(line => stripQuotes(line.replace(/\r?\n$/, ''), fence.quoteDepth).rest)
				.join('\n');
			try {
				const result = await api.queryMarkdown(source, sourcePath);
				if (result.successful && typeof result.value === 'string') {
					// A query without results is left out: an empty table or list means nothing to the reader
					if (!isEmptyQueryResult(result.value)) {
						const eol = block[0].endsWith('\r\n') ? '\r\n' : '\n';
						out.push(...result.value.split('\n').map(line => `${fence.quote}${line}${eol}`), eol);
					}
				} else {
					failures.push(/calendar/i.test(result.error ?? '') ? 'Dataview calendar query (not supported)' : 'Dataview query (error)');
				}
			} catch {
				failures.push('Dataview query (error)');
			}
		}
		return out.join('');
	}

	private expandInlineQueries(line: string, api: DataviewApi, sourcePath: string, failures: string[]): string {
		const prefix = api.settings.inlineQueryPrefix;
		const jsPrefix = api.settings.inlineJsQueryPrefix;
		return line.replace(INLINE_CODE, (span: string, ticks: string) => {
			const text = span.slice(ticks.length, span.length - ticks.length).trim();
			if (jsPrefix && text.startsWith(jsPrefix)) {
				failures.push('inline DataviewJS query (not supported)');
				return '';
			}
			if (!prefix || !text.startsWith(prefix)) return span;
			try {
				const result = api.evaluateInline(text.slice(prefix.length).trim(), sourcePath);
				if (result.successful) {
					return api.value.toString(result.value, api.settings);
				}
			} catch {
				// Reported below like any other failed query
			}
			failures.push('inline Dataview query (error)');
			return '';
		});
	}

	// Waits for every Tasks query block to finish drawing, then removes Tasks' own controls.
	private async settleTasksQueries(container: HTMLElement, failures: string[]): Promise<void> {
		// A tasks code block nobody took over means the Tasks plugin is missing or disabled
		for (const code of Array.from(container.querySelectorAll('pre > code.language-tasks'))) {
			code.parentElement?.remove();
			failures.push('Tasks query (Tasks not available)');
		}

		const pending = new Map<HTMLElement, { html: string; stablePolls: number }>();
		for (const block of Array.from(container.querySelectorAll<HTMLElement>('.block-language-tasks'))) {
			pending.set(block, { html: '', stablePolls: 0 });
		}
		const deadline = Date.now() + TASKS_TIMEOUT_MS;
		while (pending.size > 0 && Date.now() < deadline) {
			await sleep(TASKS_POLL_MS);
			for (const [block, state] of Array.from(pending.entries())) {
				const status = this.tasksBlockStatus(block, state);
				if (status === 'pending') continue;
				pending.delete(block);
				if (status === 'error') {
					block.remove();
					failures.push('Tasks query (error)');
				} else {
					this.cleanTasksResults(block);
					// A query without results is left out
					if (!block.querySelector('li')) block.remove();
				}
			}
		}
		for (const block of Array.from(pending.keys())) {
			block.remove();
			failures.push('Tasks query (timed out)');
		}
	}

	private tasksBlockStatus(block: HTMLElement, state: { html: string; stablePolls: number }): 'pending' | 'done' | 'error' {
		const isError = Array.from(block.querySelectorAll('pre'))
			.some(pre => !pre.classList.contains('plugin-tasks-query-explanation') && (pre.textContent ?? '').startsWith('Tasks query:'));
		if (isError) return 'error';
		if (block.childElementCount === 0 || (block.textContent ?? '').trim() === 'Loading Tasks ...') {
			state.stablePolls = 0;
			return 'pending';
		}
		// The task count is drawn last; without it (hidden by the query), wait until nothing changes
		if (block.querySelector('.task-count')) return 'done';
		const html = block.innerHTML;
		if (html === state.html) {
			state.stablePolls++;
		} else {
			state.html = html;
			state.stablePolls = 0;
		}
		return state.stablePolls >= TASKS_STABLE_POLLS ? 'done' : 'pending';
	}

	private cleanTasksResults(block: HTMLElement): void {
		const controls = '.plugin-tasks-toolbar, .plugin-tasks-query-explanation, .task-count, .tasks-group-count, .tasks-backlink, .tasks-edit, .tasks-postpone';
		for (const el of Array.from(block.querySelectorAll(controls))) {
			el.remove();
		}
		// Tags hidden by the query are only hidden with CSS, which an email doesn't have
		for (const tag of Array.from(block.querySelectorAll('.tasks-layout-hide-tags a.tag'))) {
			tag.remove();
		}
	}

	// Email clients don't display checkboxes: done tasks become ☑, any other status ☐.
	private convertCheckboxes(container: HTMLElement): void {
		for (const input of Array.from(container.querySelectorAll<HTMLInputElement>('input[type="checkbox"]'))) {
			const status = input.closest('li')?.getAttribute('data-task');
			const done = status !== null && status !== undefined ? status.toLowerCase() === 'x' : input.checked;
			input.replaceWith(document.createTextNode(done ? '☑ ' : '☐ '));
		}
	}

	private summarizeFailures(failures: string[]): string {
		const counts = new Map<string, number>();
		for (const failure of failures) {
			counts.set(failure, (counts.get(failure) ?? 0) + 1);
		}
		return Array.from(counts.entries()).map(([failure, count]) => `${count} ${failure}`).join(', ');
	}

	private removeInternalLinks(htmlElement: HTMLElement): void {
		for (const link of Array.from(htmlElement.querySelectorAll('a.internal-link'))) {
			link.replaceWith(document.createTextNode(link.textContent ?? ''));
		}
	}

	private async attachImages(htmlElement: HTMLElement, sourcePath: string): Promise<ImageAttachment[]> {
		const images = htmlElement.querySelectorAll('img');
		const vault = this.app.vault;
		const vaultPath = vault.adapter instanceof FileSystemAdapter ? vault.adapter.getBasePath() : '';
		const attachmentsByPath = new Map<string, ImageAttachment>();

		for (const img of Array.from(images)) {
			const src = img.getAttribute('src');
			if (!src || src.startsWith('data:')) continue;

			try {
				let imageFile: TFile | null = null;

				// Handle app:// protocol URLs (Obsidian's internal format)
				if (src.startsWith('app://')) {
					// Extract: app://vault-id/C:/path/to/vault/Daily%20Notes/attachments/file.png?timestamp
					// Remove query string and decode
					const urlWithoutQuery = src.split('?')[0];
					const match = urlWithoutQuery.match(/app:\/\/[^/]+\/(.+)/);

					if (match) {
						const fullPath = decodeURIComponent(match[1]);

						// Convert absolute path to vault-relative path
						// fullPath is like: C:/Users/.../Vault/Daily Notes/attachments/file.png
						// vaultPath is like: C:\Users\...\Vault
						const normalizedVaultPath = vaultPath.replace(/\\/g, '/');
						const normalizedFullPath = fullPath.replace(/\\/g, '/');

						if (normalizedVaultPath && normalizedFullPath.startsWith(normalizedVaultPath)) {
							// Remove vault base path to get relative path
							let relativePath = normalizedFullPath.substring(normalizedVaultPath.length);
							if (relativePath.startsWith('/')) relativePath = relativePath.substring(1);

							const abstractFile = vault.getAbstractFileByPath(relativePath);
							if (abstractFile instanceof TFile) {
								imageFile = abstractFile;
							}
						}
					}
				}
				// Handle relative paths
				else if (!src.startsWith('http')) {
					const sourceFolder = sourcePath.substring(0, sourcePath.lastIndexOf('/'));
					const imagePath = normalizePath(path.join(sourceFolder, src));
					const abstractFile = vault.getAbstractFileByPath(imagePath);
					if (abstractFile instanceof TFile) {
						imageFile = abstractFile;
					}
				}
				// Skip external URLs
				else {
					continue;
				}

				if (imageFile && imageFile instanceof TFile) {
					// The same image used twice is attached once
					let attachment = attachmentsByPath.get(imageFile.path);
					if (!attachment) {
						attachment = {
							cid: `image${attachmentsByPath.size + 1}.${Date.now()}@mail-for-outlook`,
							filename: imageFile.name,
							mimeType: this.getMimeType(imageFile.extension.toLowerCase()),
							data: await vault.readBinary(imageFile)
						};
						attachmentsByPath.set(imageFile.path, attachment);
					}
					img.setAttribute('src', `cid:${attachment.cid}`);
				}
			} catch (error) {
				console.error('Failed to attach image:', src, error);
			}
		}

		return Array.from(attachmentsByPath.values());
	}

	// The clipboard has no attachments: each image found by attachImages is written into the HTML itself.
	private embedImages(htmlElement: HTMLElement, attachments: ImageAttachment[]): void {
		const dataUris = new Map(attachments.map(attachment => [
			`cid:${attachment.cid}`,
			`data:${attachment.mimeType};base64,${Buffer.from(attachment.data).toString('base64')}`
		]));
		for (const img of Array.from(htmlElement.querySelectorAll('img'))) {
			const dataUri = dataUris.get(img.getAttribute('src') ?? '');
			if (dataUri) img.setAttribute('src', dataUri);
		}
	}

	// Recipients from linked person notes, then from addresses written in the properties and the text.
	// The note's properties always count; the rest comes from `source`.
	private collectRecipients(file: TFile, source: PreparedContent): Recipient[] {
		const metadataCache = this.app.metadataCache;
		const cache = metadataCache.getFileCache(file);
		const ownAddress = this.extractAddress(this.settings.defaultFrom);
		const seenAddresses = new Set<string>();
		const recipients: Recipient[] = [];
		const add = (value: string, name: string) => {
			const address = value.trim().replace(/^mailto:/i, '');
			const key = address.toLowerCase();
			if (!EMAIL_PATTERN.test(address) || key === ownAddress || seenAddresses.has(key)) return;
			seenAddresses.add(key);
			recipients.push({ name, address });
		};

		if (this.settings.recipientsFromPersonNotes) {
			const personType = this.normalizeValue(this.settings.personType);
			const links: LinkRef[] = [
				...(cache?.frontmatterLinks ?? []).map(link => ({ link: link.link, sourcePath: file.path })),
				...source.links
			];
			const seenNotes = new Set<string>();

			for (const { link, sourcePath } of links) {
				const target = metadataCache.getFirstLinkpathDest(getLinkpath(link), sourcePath);
				if (!target || target.extension !== 'md' || seenNotes.has(target.path)) continue;
				seenNotes.add(target.path);

				const frontmatter = metadataCache.getFileCache(target)?.frontmatter;
				if (!frontmatter) continue;

				const types = this.toStringList(this.getProperty(frontmatter, 'type')).map(value => this.normalizeValue(value));
				if (!personType || !types.includes(personType)) continue;

				for (const value of this.toStringList(this.getProperty(frontmatter, this.settings.emailProperty))) {
					add(value, target.basename);
				}
			}
		}

		if (this.settings.recipientsFromText) {
			for (const address of [...extractAddressesFromValue(cache?.frontmatter), ...extractAddresses(source.markdown)]) {
				add(address, '');
			}
		}

		return recipients;
	}

	// Property names are matched case-insensitively ("Email" and "email" both work).
	private getProperty(frontmatter: Record<string, unknown>, name: string): unknown {
		const wanted = name.trim().toLowerCase();
		const key = Object.keys(frontmatter).find(k => k.toLowerCase() === wanted);
		return key === undefined ? undefined : frontmatter[key];
	}

	private toStringList(value: unknown): string[] {
		const values = Array.isArray(value) ? value : [value];
		return values.filter((v): v is string => typeof v === 'string' && v.trim() !== '');
	}

	// Compares type values case-insensitively, also when written as a link ("[[Personne]]").
	private normalizeValue(value: string): string {
		return value.trim().replace(/^\[\[|\]\]$/g, '').trim().toLowerCase();
	}

	private extractAddress(from: string): string {
		const bracketed = from.match(/<([^>]+)>/);
		return (bracketed ? bracketed[1] : from).trim().toLowerCase();
	}

	private getMimeType(extension: string): string {
		const mimeTypes: { [key: string]: string } = {
			'png': 'image/png',
			'jpg': 'image/jpeg',
			'jpeg': 'image/jpeg',
			'gif': 'image/gif',
			'svg': 'image/svg+xml',
			'webp': 'image/webp'
		};
		return mimeTypes[extension] || 'image/png';
	}

	private wrapWithStyles(html: string): string {
		return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
      line-height: 1.6;
      color: #333;
      max-width: 800px;
      margin: 0 auto;
      padding: 20px;
    }
    img {
      max-width: 100%;
      height: auto;
    }
    code {
      background-color: #f4f4f4;
      padding: 2px 6px;
      border-radius: 3px;
      font-family: 'Courier New', Courier, monospace;
    }
    pre {
      background-color: #f4f4f4;
      padding: 10px;
      border-radius: 5px;
      overflow-x: auto;
    }
    blockquote {
      border-left: 4px solid #ddd;
      margin: 0;
      padding-left: 16px;
      color: #666;
    }
    ul.contains-task-list {
      list-style: none;
      padding-left: 1.2em;
    }
    table {
      border-collapse: collapse;
    }
    th, td {
      border: 1px solid #ddd;
      padding: 4px 8px;
    }
  </style>
</head>
<body>
  ${html}
</body>
</html>`;
	}

	// Uses Obsidian's own frontmatter parser, so Windows (CRLF) line endings are handled too.
	private removeFrontmatter(content: string): string {
		const info = getFrontMatterInfo(content);
		return info.exists ? content.slice(info.contentStart) : content;
	}

	private async createEML(subject: string, htmlContent: string, attachments: ImageAttachment[], recipients: Recipient[]): Promise<string> {
		// Build EML file manually (RFC 822 format)
		const uniqueId = `${Date.now()}_${Math.random().toString(36).slice(2, 11)}`;
		const altBoundary = `----=_Alt_${uniqueId}`;
		const relatedBoundary = `----=_Rel_${uniqueId}`;
		const from = this.settings.defaultFrom || '';
		// The safety placeholder is not an address, so the mail client refuses to send until it is removed
		const placeholder = this.settings.safetyPlaceholderEnabled ? sanitizePlaceholder(this.settings.safetyPlaceholder) : '';
		const to = [
			...(placeholder ? [placeholder] : []),
			...recipients.map(r => {
				if (!r.name) return r.address;
				const name = ASCII_PATTERN.test(r.name)
					? `"${r.name.replace(/["\\]/g, '\\$&')}"`
					: this.encodeHeaderWord(r.name);
				return `${name} <${r.address}>`;
			})
		].join(',\r\n ');

		const alternativePart = [
			`Content-Type: multipart/alternative; boundary="${altBoundary}"`,
			``,
			`--${altBoundary}`,
			`Content-Type: text/plain; charset="utf-8"`,
			`Content-Transfer-Encoding: 7bit`,
			``,
			`Please view this email in an HTML-capable email client.`,
			``,
			`--${altBoundary}`,
			`Content-Type: text/html; charset="utf-8"`,
			`Content-Transfer-Encoding: base64`,
			``,
			this.toBase64Lines(Buffer.from(htmlContent, 'utf8')),
			``,
			`--${altBoundary}--`
		];

		// Images go in a multipart/related wrapper so the HTML can reference them by Content-ID
		const body = attachments.length === 0
			? alternativePart
			: [
				`Content-Type: multipart/related; boundary="${relatedBoundary}"`,
				``,
				`--${relatedBoundary}`,
				...alternativePart,
				...attachments.flatMap(attachment => {
					const filename = this.encodeHeaderWord(attachment.filename, false);
					return [
						``,
						`--${relatedBoundary}`,
						`Content-Type: ${attachment.mimeType}; name="${filename}"`,
						`Content-Transfer-Encoding: base64`,
						`Content-ID: <${attachment.cid}>`,
						`Content-Disposition: inline; filename="${filename}"`,
						``,
						this.toBase64Lines(Buffer.from(attachment.data))
					];
				}),
				``,
				`--${relatedBoundary}--`
			];

		const emlContent = [
			`X-Unsent: 1`,
			`From: ${from}`,
			`To: ${to}`,
			`Subject: ${this.encodeHeaderWord(subject)}`,
			`MIME-Version: 1.0`,
			...body
		].join('\r\n');

		const sanitizedSubject = subject.replace(/[^a-zA-Z0-9-_]/g, '_');
		await fs.promises.mkdir(TEMP_FOLDER, { recursive: true });
		const tempPath = path.join(TEMP_FOLDER, `${sanitizedSubject}_${uniqueId}.eml`);
		await fs.promises.writeFile(tempPath, emlContent, 'utf8');
		return tempPath;
	}

	// RFC 2047 encoded-words, so accented subjects, names and filenames display correctly.
	// Header text is split into words of at most 75 characters without cutting a character in two;
	// a filename parameter can't be folded, so it stays one word.
	private encodeHeaderWord(text: string, fold = true): string {
		if (ASCII_PATTERN.test(text)) return text;
		const chunks: string[] = [];
		let chunk = '';
		for (const char of Array.from(text)) {
			if (fold && chunk && Buffer.byteLength(chunk + char, 'utf8') > 45) {
				chunks.push(chunk);
				chunk = '';
			}
			chunk += char;
		}
		chunks.push(chunk);
		return chunks.map(c => `=?UTF-8?B?${Buffer.from(c, 'utf8').toString('base64')}?=`).join('\r\n ');
	}

	private toBase64Lines(data: Buffer): string {
		return data.toString('base64').replace(/.{1,76}/g, '$&\r\n').replace(/\r\n$/, '');
	}

	private scheduleTempCleanup(): void {
		const timer = window.setTimeout(() => void this.cleanupTempFiles(), TEMP_FILE_MAX_AGE_MS + 1000);
		this.register(() => window.clearTimeout(timer));
	}

	// Deletes exported drafts older than TEMP_FILE_MAX_AGE_MS from the plugin's temp subfolder.
	private async cleanupTempFiles(): Promise<void> {
		try {
			const names = await fs.promises.readdir(TEMP_FOLDER);
			const now = Date.now();
			for (const name of names) {
				if (!name.endsWith('.eml')) continue;
				const filePath = path.join(TEMP_FOLDER, name);
				const stats = await fs.promises.stat(filePath);
				if (now - stats.mtimeMs >= TEMP_FILE_MAX_AGE_MS) {
					await fs.promises.unlink(filePath);
				}
			}
		} catch {
			// Folder not created yet, or a file is still open in the mail client: retry on the next run.
		}
	}

	// -----------------------------------------------------------------------------------------
	// Import (1.4.0)
	// -----------------------------------------------------------------------------------------

	// The Import email command: through the Templater template when possible, otherwise the built-in fallback.
	async importEmail(): Promise<void> {
		const templateFile = this.findVaultFile(this.settings.emailTemplate);
		const createNote = this.getTemplaterCreate();
		if (templateFile && createNote) {
			try {
				await createNote(templateFile, undefined, undefined, true);
			} catch (error) {
				// A half-made note may exist: no fallback, so there is never a second note
				new Notice(`❌ Import failed: ${error instanceof Error ? error.message : String(error)}`, NOTICE_ERROR_MS);
				console.error('Mail for Outlook import error:', error);
			}
			return;
		}
		new Notice('Templater indisponible — modèle de secours utilisé');
		await this.importWithFallback();
	}

	// Templater's create_new_note_from_template, if Templater is enabled and still has it.
	private getTemplaterCreate(): CreateNoteFromTemplate | null {
		const plugins = (this.app as unknown as { plugins?: { plugins?: Record<string, unknown> } }).plugins?.plugins;
		const templater = (plugins?.[TEMPLATER_ID] as { templater?: TemplaterInternals } | undefined)?.templater;
		if (!templater || typeof templater.create_new_note_from_template !== 'function') return null;
		const internals = templater as { create_new_note_from_template: CreateNoteFromTemplate };
		return (template, folder, filename, openNewNote) => internals.create_new_note_from_template(template, folder, filename, openNewNote);
	}

	// The file at a path from the settings; ".md" may be left out.
	private findVaultFile(setting: string): TFile | null {
		const trimmed = setting.trim();
		if (!trimmed) return null;
		const vaultPath = normalizePath(trimmed);
		for (const candidate of [vaultPath, `${vaultPath}.md`]) {
			const file = this.app.vault.getAbstractFileByPath(candidate);
			if (file instanceof TFile) return file;
		}
		return null;
	}

	// Import without Templater: the fallback token template, filled and saved by the plugin.
	private async importWithFallback(): Promise<void> {
		try {
			const email = await this.readEmailFromClipboard();
			const answer = await new TitleModal(this.app, 'Titre de l\'email (Objet) :', email ? sanitizeTitle(email.subject) : '').openAndWait();
			if (answer === null) {
				new Notice('⛔ Import annulé');
				return;
			}
			const title = sanitizeTitle(answer) || UNTITLED;
			const folder = normalizeFolder(this.settings.importFolder);
			const archived = this.findArchivedName({
				folder,
				baseName: this.settings.importPrefix + title,
				sentProperty: this.settings.sentProperty,
				sent: email ? email.sent : null
			});
			if (archived.duplicate) {
				const importAnyway = await new ChoiceModal(this.app, `Cet e-mail est déjà archivé : ${archived.duplicate.basename}`, '⛔ Annuler', '✅ Importer quand même').openAndWait();
				if (!importAnyway) {
					new Notice('⛔ Import annulé');
					return;
				}
			}

			const template = await this.readFallbackTemplate();
			const frontmatter = getFrontMatterInfo(template);
			const content = fillTemplate(template, frontmatter.exists ? frontmatter.contentStart : 0, archived.name, email, format => moment().format(format));

			if (folder && !this.app.vault.getAbstractFileByPath(folder)) {
				await this.app.vault.createFolder(folder);
			}
			const file = await this.app.vault.create(normalizePath(`${folder}/${archived.name}.md`), content);
			await this.app.workspace.getLeaf(false).openFile(file);
		} catch (error) {
			new Notice(`❌ Import failed: ${error instanceof Error ? error.message : String(error)}`, NOTICE_ERROR_MS);
			console.error('Mail for Outlook import error:', error);
		}
	}

	// The fallback template file, or the built-in default when none is set or it can't be found.
	private async readFallbackTemplate(): Promise<string> {
		if (!this.settings.fallbackTemplate.trim()) return DEFAULT_FALLBACK_TEMPLATE;
		const file = this.findVaultFile(this.settings.fallbackTemplate);
		if (!file) {
			new Notice(`⚠️ Fallback template not found: ${this.settings.fallbackTemplate} — built-in default used`, NOTICE_WARNING_MS);
			return DEFAULT_FALLBACK_TEMPLATE;
		}
		return this.app.vault.read(file);
	}

	// API: the email on the clipboard with its people matched to person notes, or null when the
	// clipboard holds no Outlook header (or can't be read).
	async readEmailFromClipboard(): Promise<ParsedEmail | null> {
		let markdown: string;
		try {
			markdown = await this.readClipboardAsMarkdown();
		} catch (error) {
			new Notice(`❌ Couldn't read the clipboard: ${error instanceof Error ? error.message : String(error)}`, NOTICE_ERROR_MS);
			console.error('Mail for Outlook clipboard error:', error);
			return null;
		}
		const parsed = parseEmail(markdown);
		if (!parsed) return null;

		const links = await this.linkPeople([...parsed.from, ...parsed.to, ...parsed.cc]);
		const toEmailPeople = (people: RawPerson[]): EmailPerson[] => {
			const seen = new Set<string>();
			const result: EmailPerson[] = [];
			for (const person of people) {
				const link = links.get(personKey(person)) ?? fallbackLink(person);
				// The same note appearing twice in one list is kept once
				if (seen.has(link)) continue;
				seen.add(link);
				result.push({ name: person.name, address: person.address, link });
			}
			return result;
		};
		return {
			subject: parsed.subject,
			sent: parsed.sent,
			sentRaw: parsed.sentRaw,
			from: toEmailPeople(parsed.from),
			to: toEmailPeople(parsed.to),
			cc: toEmailPeople(parsed.cc),
			header: parsed.header,
			body: parsed.body
		};
	}

	// HTML when the clipboard has it (images removed, then converted to markdown), otherwise plain text.
	private async readClipboardAsMarkdown(): Promise<string> {
		const items = await navigator.clipboard.read();
		const htmlItem = items.find(item => item.types.includes('text/html'));
		if (htmlItem) {
			const html = await (await htmlItem.getType('text/html')).text();
			const doc = new DOMParser().parseFromString(html, 'text/html');
			// Images are not imported, so no attachment files are created; style and script text is not content
			for (const element of Array.from(doc.querySelectorAll('img, style, script, title'))) {
				element.remove();
			}
			return htmlToMarkdown(doc.body);
		}
		const textItem = items.find(item => item.types.includes('text/plain'));
		return textItem ? (await textItem.getType('text/plain')).text() : '';
	}

	// Links for each person, keyed by personKey: by address, then exact name, then closest names.
	// People matched by name go through the review window first.
	private async linkPeople(people: RawPerson[]): Promise<Map<string, string>> {
		const links = new Map<string, string>();
		const rows: ReviewRow[] = [];
		const notes = this.buildPersonIndex();

		for (const person of people) {
			const key = personKey(person);
			if (links.has(key) || rows.some(row => row.key === key)) continue;

			// 1. By address: linked directly
			if (person.address) {
				const address = person.address.toLowerCase();
				const note = notes.find(n => n.addresses.has(address));
				if (note) {
					links.set(key, `[[${note.file.basename}]]`);
					continue;
				}
			}

			// 2. By name, as an unordered set of words; 3. closest names if nothing matched
			const words = nameWords(person.name);
			if (words.length > 0) {
				const exact = notes.filter(n => n.names.some(names => sameWords(names, words)));
				if (exact.length > 0) {
					rows.push({ key, person, candidates: exact.map(n => n.file), exact: true });
					continue;
				}
				const close = notes
					.map(n => ({ file: n.file, score: Math.max(0, ...n.names.map(names => wordSetSimilarity(names, words))) }))
					.filter(c => c.score >= NAME_MATCH_THRESHOLD)
					.sort((a, b) => b.score - a.score)
					.slice(0, MAX_NAME_CANDIDATES);
				if (close.length > 0) {
					rows.push({ key, person, candidates: close.map(c => c.file), exact: false });
					continue;
				}
			}

			// No candidate at all (e.g. an external contact)
			links.set(key, fallbackLink(person));
		}

		if (rows.length === 0) return links;

		const choices = await new PeopleReviewModal(this.app, rows).openAndWait();
		const changed = new Set<string>();
		for (const row of rows) {
			const choice = choices?.get(row.key);
			if (!choice || !choice.file) {
				// Skip, Escape or "No note"
				links.set(row.key, fallbackLink(row.person));
				continue;
			}
			links.set(row.key, `[[${choice.file.basename}]]`);
			if (choice.addAddress && row.person.address && await this.addAddressToNote(choice.file, row.person.address)) {
				changed.add(choice.file.path);
			}
		}
		if (changed.size > 0) {
			new Notice(`Address added to ${changed.size} person note${changed.size === 1 ? '' : 's'}`);
		}
		return links;
	}

	// Person notes (same rules as recipients), read from the metadata cache only.
	private buildPersonIndex(): PersonNote[] {
		const personType = this.normalizeValue(this.settings.personType);
		const notes: PersonNote[] = [];
		if (!personType) return notes;
		for (const file of this.app.vault.getMarkdownFiles()) {
			const frontmatter = this.app.metadataCache.getFileCache(file)?.frontmatter;
			if (!frontmatter) continue;
			const types = this.toStringList(this.getProperty(frontmatter, 'type')).map(value => this.normalizeValue(value));
			if (!types.includes(personType)) continue;
			const addresses = new Set(this.toStringList(this.getProperty(frontmatter, this.settings.emailProperty))
				.map(normalizeAddressValue)
				.filter(address => address !== ''));
			const names = [file.basename, ...(parseFrontMatterAliases(frontmatter) ?? [])]
				.map(nameWords)
				.filter(words => words.length > 0);
			notes.push({ file, addresses, names });
		}
		return notes;
	}

	// Adds an address to a person note's email property; returns whether the note changed.
	// Other values and properties are never changed.
	private async addAddressToNote(file: TFile, address: string): Promise<boolean> {
		const wanted = this.settings.emailProperty.trim();
		if (!wanted) return false;
		let added = false;
		await this.app.fileManager.processFrontMatter(file, (frontmatter: Record<string, unknown>) => {
			const key = Object.keys(frontmatter).find(k => k.toLowerCase() === wanted.toLowerCase()) ?? wanted;
			const value = frontmatter[key];
			const lower = address.toLowerCase();
			if (value === undefined || value === null || value === '') {
				frontmatter[key] = [address];
				added = true;
			} else if (typeof value === 'string') {
				if (normalizeAddressValue(value) !== lower) {
					frontmatter[key] = [value, address];
					added = true;
				}
			} else if (Array.isArray(value)) {
				if (!value.some(item => typeof item === 'string' && normalizeAddressValue(item) === lower)) {
					value.push(address);
					added = true;
				}
			}
		});
		return added;
	}

	// API: the name to archive an email under, and an archived note with the same sent time if any.
	findArchivedName(options: FindArchivedNameOptions): ArchivedName {
		const folderPath = normalizeFolder(String(options.folder ?? ''));
		const baseName = String(options.baseName ?? '');
		const sent = normalizeSentValue(options.sent);
		const folder = folderPath ? this.app.vault.getAbstractFileByPath(folderPath) : this.app.vault.getRoot();

		// Notes named baseName (number 1) or "baseName - N"; compared without regard to case,
		// since Windows and macOS file systems would treat them as the same file
		const existing = new Map<number, TFile>();
		if (folder instanceof TFolder) {
			for (const child of folder.children) {
				if (!(child instanceof TFile) || child.extension !== 'md') continue;
				const number = archiveNumber(child.basename, baseName);
				if (number !== null && !existing.has(number)) existing.set(number, child);
			}
		}

		let duplicate: TFile | null = null;
		if (sent) {
			for (const number of Array.from(existing.keys()).sort((a, b) => a - b)) {
				const file = existing.get(number);
				const frontmatter = file ? this.app.metadataCache.getFileCache(file)?.frontmatter : undefined;
				if (file && frontmatter && normalizeSentValue(this.getProperty(frontmatter, options.sentProperty)) === sent) {
					duplicate = file;
					break;
				}
			}
		}

		let number = 1;
		while (existing.has(number)) number = number === 1 ? 2 : number + 1;
		return { name: number === 1 ? baseName : `${baseName} - ${number}`, duplicate };
	}

	async loadSettings() {
		this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData() as Partial<MailForOutlookSettings>);
	}

	async saveSettings() {
		await this.saveData(this.settings);
	}
}

// ---------------------------------------------------------------------------------------------
// Markdown scanning helpers
// ---------------------------------------------------------------------------------------------

// Splits text into lines, each keeping its own line ending.
function splitLines(text: string): string[] {
	return text.match(/[^\n]*\n|[^\n]+$/g) ?? [];
}

function lineIndexAt(text: string, offset: number): number {
	let count = 0;
	for (let i = 0; i < offset && i < text.length; i++) {
		if (text[i] === '\n') count++;
	}
	return count;
}

// Removes up to `depth` blockquote markers ("> ") from the start of a line.
function stripQuotes(line: string, depth: number): { rest: string; removed: number } {
	let rest = line;
	let removed = 0;
	while (removed < depth) {
		const marker = rest.match(/^[ \t]{0,3}>[ ]?/);
		if (!marker) break;
		rest = rest.slice(marker[0].length);
		removed++;
	}
	return { rest, removed };
}

// A closing fence uses the same character as the opening one, at least as many times.
function isFenceClose(line: string, fence: Fence): boolean {
	const { rest } = stripQuotes(line.replace(/\r?\n$/, ''), fence.quoteDepth);
	const close = rest.match(/^ {0,3}(`{3,}|~{3,})[ \t]*$/);
	return !!close && close[1][0] === fence.marker[0] && close[1].length >= fence.marker.length;
}

// Finds fenced code blocks (``` or ~~~), including ones inside blockquotes and callouts.
function findFences(lines: string[]): Fence[] {
	const fences: Fence[] = [];
	let open: Fence | null = null;
	for (let i = 0; i < lines.length; i++) {
		const content = lines[i].replace(/\r?\n$/, '');
		if (open) {
			if (stripQuotes(content, open.quoteDepth).removed < open.quoteDepth) {
				// The blockquote around the fence ended, which closes the fence
				open.end = i - 1;
				fences.push(open);
				open = null;
			} else {
				if (isFenceClose(content, open)) {
					open.end = i;
					fences.push(open);
					open = null;
				}
				continue;
			}
		}
		const quoteMatch = content.match(/^(?:[ \t]{0,3}>[ ]?)*/);
		const quote = quoteMatch ? quoteMatch[0] : '';
		const opening = content.slice(quote.length).match(/^ {0,3}(`{3,}|~{3,})[ \t]*([^\s`]*)/);
		if (opening && !(opening[1][0] === '`' && content.slice(quote.length + opening[0].length).includes('`'))) {
			open = { start: i, end: i, marker: opening[1], lang: opening[2].toLowerCase(), quote, quoteDepth: (quote.match(/>/g) ?? []).length };
		}
	}
	if (open) {
		open.end = lines.length - 1;
		fences.push(open);
	}
	return fences;
}

function fenceLineFlags(lines: string[]): boolean[] {
	const flags = new Array<boolean>(lines.length).fill(false);
	for (const fence of findFences(lines)) {
		for (let i = fence.start; i <= fence.end; i++) flags[i] = true;
	}
	return flags;
}

// Text a reader actually sees: code, comments and fenced blocks are blanked out.
function visibleText(markdown: string): string {
	const lines = splitLines(markdown);
	const inFence = fenceLineFlags(lines);
	return lines
		.map((line, i) => inFence[i] ? '\n' : line.replace(INLINE_CODE, ' '))
		.join('')
		.replace(COMMENTS, ' ');
}

// Link targets of wikilinks and Markdown links to notes (not web or mailto links).
function extractLinks(markdown: string): string[] {
	const text = visibleText(markdown);
	const links: string[] = [];
	for (const match of Array.from(text.matchAll(/!?\[\[([^\]]+?)\]\]/g))) {
		const link = match[1].split('|')[0].trim();
		if (link) links.push(link);
	}
	for (const match of Array.from(text.matchAll(/!?\[[^\]]*\]\(([^)]+)\)/g))) {
		let target = match[1].trim().replace(/^<|>$/g, '').split(/\s+"/)[0];
		if (!target || /^[a-z][a-z0-9+.-]*:/i.test(target)) continue;
		try {
			target = decodeURI(target);
		} catch {
			// Keep the target as written
		}
		links.push(target);
	}
	return links;
}

// Email addresses written in the text, including mailto: links. Addresses inside code, comments,
// web links, file names and link targets are ignored.
function extractAddresses(markdown: string): string[] {
	const text = visibleText(markdown)
		.replace(/\[\[[^\]]*\]\]/g, ' ')
		.replace(/\]\((?!\s*<?mailto:)[^)]*\)/gi, '] ')
		.replace(/\b[a-z][a-z0-9+.-]*:\/\/\S+/gi, ' ');
	const addresses: string[] = [];
	for (const match of Array.from(text.matchAll(ADDRESS_IN_TEXT))) {
		const address = match[0].replace(/^[.'-]+|[.'-]+$/g, '');
		const domainEnd = address.slice(address.lastIndexOf('.') + 1);
		if (EMAIL_PATTERN.test(address) && /^[A-Za-z]{2,}$/.test(domainEnd)) {
			addresses.push(address);
		}
	}
	return addresses;
}

function extractAddressesFromValue(value: unknown): string[] {
	if (typeof value === 'string') return extractAddresses(value);
	if (Array.isArray(value)) return value.flatMap(item => extractAddressesFromValue(item));
	if (value && typeof value === 'object') return Object.values(value).flatMap(item => extractAddressesFromValue(item));
	return [];
}

// Applies the #private / #privateEnd rules:
// - #private on a list item hides only that item and what belongs to it: lines indented below it
//   (sub-items, continuation, code) and lines continuing it without a blank line.
// - Otherwise #private starts a private area at its line; if that line is the first one under a
//   heading, the heading is private too; on a heading line, the heading starts the area.
// - The area ends at #privateEnd (that line included), otherwise at the next heading of the same
//   or higher level than the section it started in (any heading, before the first heading).
// - A #privateEnd outside a private area is removed from the text. Tags in code are ignored.
function analyzePrivacy(text: string): PrivacyAnalysis {
	const lines = splitLines(text);
	const fences = findFences(lines);
	const fenceEnd = new Map(fences.map(fence => [fence.start, fence.end]));
	const inFence = new Array<boolean>(lines.length).fill(false);
	for (const fence of fences) {
		for (let i = fence.start; i <= fence.end; i++) inFence[i] = true;
	}
	const area = new Array<number>(lines.length).fill(0);
	let current = 0;
	let areaCount = 0;
	let areaLevel = 0;
	let headingLine = -1;
	let headingLevel = 0;
	let contentSinceHeading = false;
	// Private list item being followed: its area number, indentation and blockquote depth
	let listArea = 0;
	let listIndent = 0;
	let listQuoteDepth = 0;
	let listPrevBlank = false;
	const pendingBlanks: number[] = [];

	for (let i = 0; i < lines.length; i++) {
		const content = lines[i].replace(/\r?\n$/, '');

		if (listArea) {
			const { rest, removed } = stripQuotes(content, listQuoteDepth);
			const sameQuote = removed === listQuoteDepth;
			if (sameQuote && rest.trim() === '' && !inFence[i]) {
				// Blank lines belong to the item only if more of the item follows
				pendingBlanks.push(i);
				listPrevBlank = true;
				continue;
			}
			const isHeading = !inFence[i] && HEADING.test(content);
			const lazy = !listPrevBlank && !inFence[i] && !LIST_ITEM.test(rest);
			if (sameQuote && !isHeading && (indentWidth(rest) > listIndent || lazy)) {
				for (const j of pendingBlanks) area[j] = listArea;
				pendingBlanks.length = 0;
				const end = fenceEnd.get(i) ?? i;
				for (let j = i; j <= end; j++) area[j] = listArea;
				i = end;
				listPrevBlank = false;
				contentSinceHeading = true;
				continue;
			}
			// Anything else ends the item; this line is handled normally below
			listArea = 0;
			pendingBlanks.length = 0;
		}

		if (inFence[i]) {
			if (current) area[i] = current;
			contentSinceHeading = true;
			continue;
		}
		const heading = content.match(HEADING);
		const level = heading ? heading[1].length : 0;
		const scan = content.replace(INLINE_CODE, ' ');
		const hasStart = TAG_PRIVATE.test(scan);
		const hasEnd = TAG_PRIVATE_END.test(scan);

		if (current) {
			if (heading && (areaLevel === 0 || level <= areaLevel)) {
				current = 0;
			} else {
				area[i] = current;
				if (hasEnd) current = 0;
				continue;
			}
		}

		if (heading) {
			headingLine = i;
			headingLevel = level;
			contentSinceHeading = false;
		}

		if (hasStart && !heading) {
			const { rest, removed } = stripQuotes(content, Number.MAX_SAFE_INTEGER);
			if (LIST_ITEM.test(rest)) {
				listArea = ++areaCount;
				listIndent = indentWidth(rest);
				listQuoteDepth = removed;
				listPrevBlank = false;
				area[i] = listArea;
				contentSinceHeading = true;
				continue;
			}
		}

		if (hasStart) {
			current = ++areaCount;
			if (heading) {
				areaLevel = level;
				area[i] = current;
			} else if (headingLine >= 0 && !contentSinceHeading) {
				areaLevel = headingLevel;
				for (let j = headingLine; j <= i; j++) area[j] = current;
			} else {
				areaLevel = headingLevel;
				area[i] = current;
			}
			if (hasEnd) current = 0;
			continue;
		}

		if (hasEnd) {
			lines[i] = lines[i].replace(TAG_PRIVATE_END_ALL, '$1');
		}
		if (!heading && content.trim() !== '') {
			contentSinceHeading = true;
		}
	}
	return { lines, area };
}

// Width of the leading whitespace, tabs counting up to the next multiple of 4.
function indentWidth(line: string): number {
	let width = 0;
	for (const char of line) {
		if (char === ' ') width++;
		else if (char === '\t') width += 4 - (width % 4);
		else break;
	}
	return width;
}

// Dataview returns an empty string for an empty list, and only the header rows for an empty table.
function isEmptyQueryResult(markdown: string): boolean {
	const lines = markdown.split('\n').filter(line => line.trim() !== '');
	return lines.length === 0 || (lines.length === 2 && /^\|?[\s:|-]*-[\s:|-]*\|?$/.test(lines[1].trim()));
}

// Joins the public lines of [from, to) and counts the private areas met there.
function keepPublic(analysis: PrivacyAnalysis, from = 0, to = analysis.lines.length): { text: string; areas: number } {
	const areas = new Set<number>();
	let text = '';
	for (let i = from; i < to && i < analysis.lines.length; i++) {
		if (analysis.area[i]) {
			areas.add(analysis.area[i]);
		} else {
			text += analysis.lines[i];
		}
	}
	return { text, areas: areas.size };
}

// The placeholder must stay a single token that is clearly not an email address.
function sanitizePlaceholder(text: string): string {
	const cleaned = text.trim().replace(/\s+/g, '-').replace(/[^A-Za-z0-9!#$%&'*+/=?^_`{|}~.-]/g, '');
	return cleaned || DEFAULT_SETTINGS.safetyPlaceholder;
}

// ---------------------------------------------------------------------------------------------
// Import helpers (1.4.0)
// ---------------------------------------------------------------------------------------------

// A vault folder path from a setting or a template: leading and trailing slashes removed, "" for the root.
function normalizeFolder(folder: string): string {
	const normalized = normalizePath(folder.trim());
	return normalized === '/' ? '' : normalized;
}

// A subject made usable as a file name and inside the "# [[…]]" heading link.
function sanitizeTitle(text: string): string {
	return text.replace(TITLE_ILLEGAL, '').trim();
}

// Lowercase, without accents.
function foldText(text: string): string {
	return text.normalize('NFD').replace(/\p{M}/gu, '').toLowerCase();
}

// One markdown line made plain for header parsing; the body keeps its markdown.
function normalizeHeaderLine(line: string): string {
	return line
		.replace(/[\u00A0\u202F\u2007]/g, ' ')
		.replace(/^(?:[ \t]*>)+[ \t]?/, '')
		.replace(/\*\*|__/g, '')
		.replace(/\\([\\`*_{}[\]()#+\-.!<>|~=])/g, '$1')
		.replace(/\[([^\]]*)\]\(\s*<?mailto:([^)\s>]+)>?\s*\)/gi, '$2')
		.replace(/<\/?[A-Za-z][A-Za-z0-9-]*(?:\s[^<>]*)?\/?>/g, '')
		.trim();
}

// The label of a header line: a known field, null for another label ("Importance :"),
// undefined when the line starts with no label.
function headerLabel(line: string): HeaderField | null | undefined {
	const match = line.match(LABEL_LINE);
	if (!match) return undefined;
	return HEADER_LABELS[foldText(match[1].trim())] ?? null;
}

// The first From line followed, within the next HEADER_WINDOW non-empty lines, by a Sent line
// and then a Subject line. The block ends at the Subject line.
function findHeaderBlock(lines: string[]): { start: number; end: number } | null {
	for (let i = 0; i < lines.length; i++) {
		if (headerLabel(lines[i]) !== 'from') continue;
		let seenSent = false;
		let count = 0;
		for (let j = i + 1; j < lines.length && count < HEADER_WINDOW; j++) {
			if (lines[j] === '') continue;
			count++;
			const label = headerLabel(lines[j]);
			if (label === 'sent') {
				seenSent = true;
			} else if (label === 'subject') {
				if (seenSent) return { start: i, end: j };
				break;
			} else if (label === 'from') {
				break;
			}
		}
	}
	return null;
}

// Finds the Outlook header in clipboard markdown and reads its fields; null when there is none.
// Everything above the header (the forward draft, a signature, Outlook's separator) is discarded.
function parseEmail(markdown: string): ParsedHeader | null {
	const lines = markdown.replace(/\r\n?/g, '\n').split('\n');
	const clean = lines.map(normalizeHeaderLine);
	const block = findHeaderBlock(clean);
	if (!block) return null;

	// Fields are read only inside the block; a line with no label continues the previous field,
	// and a line with an unknown label (and what continues it) is ignored
	const fields: Partial<Record<HeaderField, string>> = {};
	let current: HeaderField | null = null;
	for (let i = block.start; i <= block.end; i++) {
		const line = clean[i];
		if (line === '') continue;
		const label = headerLabel(line);
		if (label === undefined) {
			if (current) fields[current] = `${fields[current] ?? ''} ${line}`.trim();
			continue;
		}
		if (label === null || fields[label] !== undefined) {
			current = null;
			continue;
		}
		current = label;
		fields[label] = (line.match(LABEL_LINE)?.[2] ?? '').trim();
	}

	const sentRaw = fields.sent ?? '';
	return {
		subject: fields.subject ?? '',
		sent: parseSentDate(sentRaw),
		sentRaw,
		from: parsePeople(fields.from ?? ''),
		to: parsePeople(fields.to ?? ''),
		cc: parsePeople(fields.cc ?? ''),
		header: lines.slice(block.start, block.end + 1).join('\n'),
		body: lines.slice(block.start).join('\n').trimEnd()
	};
}

// "YYYY-MM-DD HH:mm" from French long, English long or numeric (day first) dates; null if unparseable.
function parseSentDate(raw: string): string | null {
	const text = foldText(raw).replace(/,/g, ' ').replace(/\s+/g, ' ').trim();
	const time = '\\s+(?:a\\s+|at\\s+)?(\\d{1,2})[:h](\\d{2})(?:\\s*([ap])\\.?m\\.?)?';
	let day: number;
	let month: number;
	let year: number;
	let hourText: string;
	let minuteText: string;
	let meridiem: string | undefined;

	let match = text.match(new RegExp(`(\\d{1,2})[./](\\d{1,2})[./](\\d{4})${time}`));
	if (match) {
		[day, month, year] = [Number(match[1]), Number(match[2]), Number(match[3])];
		[hourText, minuteText, meridiem] = [match[4], match[5], match[6]];
	} else if ((match = text.match(new RegExp(`(\\d{1,2})(?:er)?\\.?\\s+([a-z]+)\\.?\\s+(\\d{4})${time}`))) && MONTHS[match[2]]) {
		[day, month, year] = [Number(match[1]), MONTHS[match[2]], Number(match[3])];
		[hourText, minuteText, meridiem] = [match[4], match[5], match[6]];
	} else if ((match = text.match(new RegExp(`([a-z]+)\\.?\\s+(\\d{1,2})\\s+(\\d{4})${time}`))) && MONTHS[match[1]]) {
		[day, month, year] = [Number(match[2]), MONTHS[match[1]], Number(match[3])];
		[hourText, minuteText, meridiem] = [match[4], match[5], match[6]];
	} else {
		return null;
	}

	let hour = Number(hourText);
	const minute = Number(minuteText);
	if (meridiem) {
		if (hour < 1 || hour > 12) return null;
		hour = hour % 12 + (meridiem === 'p' ? 12 : 0);
	}
	const check = new Date(year, month - 1, day);
	if (month < 1 || month > 12 || check.getDate() !== day || check.getMonth() !== month - 1 || hour > 23 || minute > 59) {
		return null;
	}
	const pad = (n: number) => String(n).padStart(2, '0');
	return `${year}-${pad(month)}-${pad(day)} ${pad(hour)}:${pad(minute)}`;
}

// People of a From / To / Cc field. Entries are split on ";"; without ";", several <…@…> are split
// after each ">" followed by "," (a comma alone is not a separator: "Lastname, Firstname").
function parsePeople(value: string): RawPerson[] {
	const text = value.trim();
	if (!text) return [];
	let entries: string[];
	if (text.includes(';')) {
		entries = text.split(';');
	} else if ((text.match(/<[^<>]*@[^<>]*>/g) ?? []).length > 1) {
		entries = text.split(/>\s*,/).map((part, i, parts) => i < parts.length - 1 ? `${part}>` : part);
	} else {
		entries = [text];
	}
	const people: RawPerson[] = [];
	for (const entry of entries) {
		const person = parsePerson(entry);
		if (person) people.push(person);
	}
	return people;
}

// Address inside <>, or a bare address, or mailto:; name = text before it, quotes removed.
function parsePerson(entry: string): RawPerson | null {
	const text = entry.trim();
	if (!text) return null;
	let name: string;
	let address: string | null = null;
	const bracketed = text.match(/<([^<>]*)>/);
	if (bracketed) {
		address = cleanAddress(bracketed[1]);
		name = text.slice(0, bracketed.index);
	} else {
		const bare = text.match(/(?:mailto:)?[^\s<>"';,]+@[^\s<>"';,]+/i);
		if (bare) {
			address = cleanAddress(bare[0]);
			name = text.slice(0, bare.index);
		} else {
			name = text;
		}
	}
	name = name.replace(/["“”«»]/g, '').replace(/^\s*'+|'+\s*$/g, '').trim();
	// Outlook sometimes repeats the address as the name
	if (address && name.toLowerCase() === address.toLowerCase()) name = '';
	if (!name && !address) return null;
	return { name, address };
}

function cleanAddress(text: string): string | null {
	const address = text.trim().replace(/^mailto:/i, '').trim();
	return EMAIL_PATTERN.test(address) ? address : null;
}

// An address from a person note's email property: "x", "mailto:x" or "[x](mailto:x)", lowercased.
function normalizeAddressValue(value: string): string {
	const link = value.match(/\]\(\s*<?mailto:([^)\s>]+)>?\s*\)/i);
	return (link ? link[1] : value).trim().replace(/^<|>$/g, '').replace(/^mailto:/i, '').trim().toLowerCase();
}

// Display name or note name as words: (…) and […] segments, quotes and commas removed, no accents,
// lowercase. Hyphenated names stay one word.
function nameWords(name: string): string[] {
	const text = foldText(name.replace(/\([^)]*\)|\[[^\]]*\]/g, ' ').replace(/["'“”‘’«»,]/g, ' '));
	return Array.from(new Set(text.split(/\s+/).filter(word => word !== '')));
}

function sameWords(a: string[], b: string[]): boolean {
	return a.length === b.length && a.every(word => b.includes(word));
}

function levenshtein(a: string, b: string): number {
	let previous = Array.from({ length: b.length + 1 }, (_, j) => j);
	for (let i = 1; i <= a.length; i++) {
		const current = [i];
		for (let j = 1; j <= b.length; j++) {
			current[j] = Math.min(previous[j] + 1, current[j - 1] + 1, previous[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
		}
		previous = current;
	}
	return previous[b.length];
}

// 1 for identical words, 0 for nothing in common.
function wordSimilarity(a: string, b: string): number {
	const length = Math.max(a.length, b.length);
	return length === 0 ? 1 : 1 - levenshtein(a, b) / length;
}

// Similarity of two names as word sets: the best pairing of their words, averaged over the longer
// name, so a missing or extra word counts as 0.
function wordSetSimilarity(a: string[], b: string[]): number {
	const [short, long] = a.length <= b.length ? [a, b] : [b, a];
	if (long.length === 0) return 0;
	const scores = short.map(x => long.map(y => wordSimilarity(x, y)));
	let best = 0;
	if (long.length <= MAX_EXHAUSTIVE_WORDS) {
		const used = new Array<boolean>(long.length).fill(false);
		const visit = (i: number, total: number) => {
			if (i === short.length) {
				best = Math.max(best, total);
				return;
			}
			for (let j = 0; j < long.length; j++) {
				if (used[j]) continue;
				used[j] = true;
				visit(i + 1, total + scores[i][j]);
				used[j] = false;
			}
		};
		visit(0, 0);
	} else {
		const used = new Set<number>();
		for (const row of scores) {
			let pick = -1;
			row.forEach((score, j) => {
				if (!used.has(j) && (pick < 0 || score > row[pick])) pick = j;
			});
			if (pick >= 0) {
				used.add(pick);
				best += row[pick];
			}
		}
	}
	return best / long.length;
}

// One key per person, so a person appearing in several fields is matched and reviewed once.
function personKey(person: RawPerson): string {
	return person.address ? `address:${person.address.toLowerCase()}` : `name:${nameWords(person.name).join(' ')}`;
}

// The link of a person with no person note: "[[address]]", or the plain name.
function fallbackLink(person: RawPerson): string {
	return person.address ? `[[${person.address}]]` : person.name;
}

// 1 for a note named baseName, N for "baseName - N" (N ≥ 2), null otherwise; case is ignored.
function archiveNumber(name: string, baseName: string): number | null {
	const lowerName = name.toLowerCase();
	const lowerBase = baseName.toLowerCase();
	if (lowerName === lowerBase) return 1;
	if (!lowerName.startsWith(`${lowerBase} - `)) return null;
	const suffix = lowerName.slice(lowerBase.length + 3);
	if (!/^\d+$/.test(suffix)) return null;
	const number = Number(suffix);
	return number >= 2 ? number : null;
}

// A sent date as "YYYY-MM-DD HH:mm", from a string or a Date property value; null otherwise.
function normalizeSentValue(value: unknown): string | null {
	if (value instanceof Date) {
		if (isNaN(value.getTime())) return null;
		// YAML timestamps without a time zone are UTC
		value = value.toISOString();
	}
	if (typeof value !== 'string') return null;
	const match = value.trim().match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})/);
	return match ? `${match[1]}-${match[2]}-${match[3]} ${match[4]}:${match[5]}` : null;
}

// Replaces the fallback template's tokens in one pass: replacement text is never parsed again,
// unknown {{…}} text and <%…%> are left as they are. Tokens before `frontmatterEnd` are written
// as YAML values: text as a double-quoted string, people as a flow list of quoted links.
function fillTemplate(template: string, frontmatterEnd: number, title: string, email: ParsedEmail | null, today: (format: string) => string): string {
	const texts: Record<string, string> = {
		title,
		subject: email ? email.subject : '',
		sent: email?.sent ?? '',
		body: email ? email.body : ''
	};
	const people: Record<string, string[]> = {
		from: email ? email.from.map(p => p.link) : [],
		to: email ? email.to.map(p => p.link) : [],
		cc: email ? email.cc.map(p => p.link) : []
	};
	return template.replace(/\{\{(title|subject|sent|from|to|cc|body|today)(?::([^{}]*))?\}\}/g, (token: string, name: string, format: string | undefined, offset: number) => {
		const inYaml = offset < frontmatterEnd;
		if (name === 'today') {
			const value = today(format?.trim() || 'YYYY-MM-DD');
			return inYaml ? JSON.stringify(value) : value;
		}
		// Only {{today}} takes a format
		if (format !== undefined) return token;
		if (name in people) {
			return inYaml ? `[${people[name].map(link => JSON.stringify(link)).join(', ')}]` : people[name].join(', ');
		}
		return inYaml ? JSON.stringify(texts[name]) : texts[name];
	});
}

// A notice that stays on screen during an export, shows its progress, and ends with the result.
class ProgressNotice {
	private notice: Notice;

	constructor(private action: string) {
		this.notice = new Notice(action, 0);
	}

	// Shows `stage` of EXPORT_STAGES as a filled bar, with what is being done.
	step(stage: number, label: string): void {
		const bar = '▰'.repeat(stage) + '▱'.repeat(EXPORT_STAGES - stage);
		this.notice.setMessage(`${this.action} ${bar} ${label}`);
	}

	// Replaces the progress with the final message, one line per entry, and closes it after `durationMs`.
	finish(lines: string[], durationMs: number): void {
		const fragment = activeDocument.createDocumentFragment();
		for (const line of lines) {
			fragment.createDiv({ text: line });
		}
		this.notice.setMessage(fragment);
		window.setTimeout(() => this.notice.hide(), durationMs);
	}
}

// Asks for the note title; resolves to null on Escape.
class TitleModal extends Modal {
	private resolve: ((value: string | null) => void) | null = null;
	private value: string;

	constructor(app: App, private prompt: string, initial: string) {
		super(app);
		this.value = initial;
	}

	openAndWait(): Promise<string | null> {
		return new Promise(resolve => {
			this.resolve = resolve;
			this.open();
		});
	}

	onOpen(): void {
		this.setTitle(this.prompt);
		new Setting(this.contentEl)
			.addText(text => {
				text.setValue(this.value).onChange(value => {
					this.value = value;
				});
				text.inputEl.addEventListener('keydown', event => {
					if (event.key === 'Enter' && !event.isComposing) {
						event.preventDefault();
						this.finish(this.value);
					}
				});
				window.setTimeout(() => text.inputEl.select(), 0);
			})
			.addButton(button => button
				.setButtonText('OK')
				.setCta()
				.onClick(() => this.finish(this.value)));
	}

	onClose(): void {
		this.contentEl.empty();
		this.finish(null);
	}

	private finish(value: string | null): void {
		const resolve = this.resolve;
		this.resolve = null;
		if (!resolve) return;
		resolve(value);
		this.close();
	}
}

// A message with two choices; resolves to true for the second one, false for the first one or Escape.
class ChoiceModal extends Modal {
	private resolve: ((value: boolean) => void) | null = null;

	constructor(app: App, private message: string, private cancelText: string, private confirmText: string) {
		super(app);
	}

	openAndWait(): Promise<boolean> {
		return new Promise(resolve => {
			this.resolve = resolve;
			this.open();
		});
	}

	onOpen(): void {
		this.contentEl.createEl('p', { text: this.message });
		new Setting(this.contentEl)
			.addButton(button => button
				.setButtonText(this.cancelText)
				.onClick(() => this.finish(false)))
			.addButton(button => button
				.setButtonText(this.confirmText)
				.onClick(() => this.finish(true)));
	}

	onClose(): void {
		this.contentEl.empty();
		this.finish(false);
	}

	private finish(value: boolean): void {
		const resolve = this.resolve;
		this.resolve = null;
		if (!resolve) return;
		resolve(value);
		this.close();
	}
}

// One window per import, for the people matched by name: choose their note, and whether to add
// their address to it. Resolves to the choices on Confirm, null on Skip or Escape.
class PeopleReviewModal extends Modal {
	private resolve: ((value: Map<string, ReviewChoice> | null) => void) | null = null;
	private choices = new Map<string, ReviewChoice>();

	constructor(app: App, private rows: ReviewRow[]) {
		super(app);
	}

	openAndWait(): Promise<Map<string, ReviewChoice> | null> {
		return new Promise(resolve => {
			this.resolve = resolve;
			this.open();
		});
	}

	onOpen(): void {
		const { contentEl } = this;
		this.setTitle('Link people to person notes');
		contentEl.createEl('p', {
			cls: 'setting-item-description',
			text: 'These people were matched by name. Choose the note of each person; without a note, the address is kept.'
		});

		for (const row of this.rows) {
			// An exact name match is likely right; a closest name less so
			const choice: ReviewChoice = { file: row.candidates[0] ?? null, addAddress: row.exact };
			this.choices.set(row.key, choice);
			let toggle: ToggleComponent | null = null;

			new Setting(contentEl)
				.setName(row.person.name || row.person.address || '')
				.setDesc(row.person.address ?? 'No address shown by Outlook')
				.addDropdown(dropdown => {
					row.candidates.forEach((file, index) => {
						// Two notes with the same name are told apart by their path
						const twin = row.candidates.some(other => other !== file && other.basename === file.basename);
						dropdown.addOption(String(index), twin ? file.path : file.basename);
					});
					dropdown.addOption('', 'No note');
					dropdown.setValue(row.candidates.length > 0 ? '0' : '');
					dropdown.onChange(value => {
						choice.file = value === '' ? null : row.candidates[Number(value)] ?? null;
						toggle?.setDisabled(choice.file === null);
					});
				});

			// No address, nothing to add
			if (row.person.address) {
				new Setting(contentEl)
					.setName('Add this address to the note')
					.addToggle(component => {
						toggle = component;
						component.setValue(choice.addAddress).onChange(value => {
							choice.addAddress = value;
						});
					});
			}
		}

		new Setting(contentEl)
			.addButton(button => button
				.setButtonText('Skip')
				.onClick(() => this.finish(null)))
			.addButton(button => button
				.setButtonText('Confirm')
				.setCta()
				.onClick(() => this.finish(this.choices)));
	}

	onClose(): void {
		this.contentEl.empty();
		this.finish(null);
	}

	private finish(value: Map<string, ReviewChoice> | null): void {
		const resolve = this.resolve;
		this.resolve = null;
		if (!resolve) return;
		resolve(value);
		this.close();
	}
}

class MailForOutlookSettingTab extends PluginSettingTab {
	plugin: MailForOutlookPlugin;

	constructor(app: App, plugin: MailForOutlookPlugin) {
		super(app, plugin);
		this.plugin = plugin;
	}

	display(): void {
		const { containerEl } = this;
		containerEl.empty();

		containerEl.createEl('p', {
			cls: 'setting-item-description',
			text: 'Send a note, or part of it, as an Outlook email draft, or copy it as rich text to paste into OneNote, Jira or other apps. Parts tagged #private are always left out.'
		});

		// Content: applies to emails and copies alike
		new Setting(containerEl).setName('Content').setHeading();

		new Setting(containerEl)
			.setName('Include frontmatter')
			.setDesc('Include the note\'s YAML frontmatter in emails and copies.')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.includeFrontmatter)
				.onChange(async (value) => {
					this.plugin.settings.includeFrontmatter = value;
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName('Expand queries')
			.setDesc('Replace Dataview and Tasks queries with their results. When off, queries are kept as code.')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.expandQueries)
				.onChange(async (value) => {
					this.plugin.settings.expandQueries = value;
					await this.plugin.saveSettings();
				}));

		// Email: settings that only matter for email drafts
		new Setting(containerEl).setName('Email').setHeading();

		new Setting(containerEl)
			.setName('Default from address')
			.setDesc('Default email address used as the sender.')
			.addText(text => text
				.setPlaceholder('you@example.com')
				.setValue(this.plugin.settings.defaultFrom)
				.onChange(async (value) => {
					this.plugin.settings.defaultFrom = value;
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName('Safety against early sending')
			.setDesc('Put a placeholder that is not an email address first among the recipients, so the email can\'t be sent until you delete it.')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.safetyPlaceholderEnabled)
				.onChange(async (value) => {
					this.plugin.settings.safetyPlaceholderEnabled = value;
					await this.plugin.saveSettings();
					// Show or hide the placeholder text
					this.display();
				}));

		if (this.plugin.settings.safetyPlaceholderEnabled) {
			new Setting(containerEl)
				.setName('Safety placeholder')
				.setDesc('Text used as the placeholder. Spaces become dashes; @ and punctuation that would make it an address are removed.')
				.addText(text => text
					.setPlaceholder(DEFAULT_SETTINGS.safetyPlaceholder)
					.setValue(this.plugin.settings.safetyPlaceholder)
					.onChange(async (value) => {
						this.plugin.settings.safetyPlaceholder = value;
						await this.plugin.saveSettings();
					}));
		}

		// Recipients: email drafts only
		new Setting(containerEl)
			.setName('Recipients')
			.setDesc('Who is added as a recipient of an email draft. Copies have no recipients.')
			.setHeading();

		new Setting(containerEl)
			.setName('From linked person notes')
			.setDesc('Add the email address of every person note linked in the note or its properties.')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.recipientsFromPersonNotes)
				.onChange(async (value) => {
					this.plugin.settings.recipientsFromPersonNotes = value;
					await this.plugin.saveSettings();
				}));

		// Always shown: Import email uses them too, even when person notes are not used for recipients
		new Setting(containerEl)
			.setName('Person note type')
			.setDesc('Notes whose "type" property has this value are treated as people: for recipients, and to link people when importing an email.')
			.addText(text => text
				.setPlaceholder(DEFAULT_SETTINGS.personType)
				.setValue(this.plugin.settings.personType)
				.onChange(async (value) => {
					this.plugin.settings.personType = value;
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName('Email property')
			.setDesc('Property in person notes that holds their email addresses: for recipients, and to match people when importing an email. Addresses you confirm on import are added to it.')
			.addText(text => text
				.setPlaceholder(DEFAULT_SETTINGS.emailProperty)
				.setValue(this.plugin.settings.emailProperty)
				.onChange(async (value) => {
					this.plugin.settings.emailProperty = value;
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName('From addresses in the text')
			.setDesc('Add every email address written in the note or its properties.')
			.addToggle(toggle => toggle
				.setValue(this.plugin.settings.recipientsFromText)
				.onChange(async (value) => {
					this.plugin.settings.recipientsFromText = value;
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName('Recipients when sending a selection')
			.setDesc('Where recipients come from when you send a selection. The note\'s properties always count.')
			.addDropdown(dropdown => dropdown
				.addOption('selection', 'Selection only')
				.addOption('note', 'Whole note')
				.setValue(this.plugin.settings.selectionRecipientScope)
				.onChange(async (value) => {
					this.plugin.settings.selectionRecipientScope = value === 'note' ? 'note' : 'selection';
					await this.plugin.saveSettings();
				}));

		// Import: the Import email command
		new Setting(containerEl)
			.setName('Import')
			.setDesc('The Import email command turns an email copied from Outlook into a note. People are linked to person notes using the two settings above.')
			.setHeading();

		new Setting(containerEl)
			.setName('Email template (Templater)')
			.setDesc('Path of the Templater template that creates the note. Empty, or Templater not enabled: the fallback template is used.')
			.addText(text => text
				.setPlaceholder('Templates/Email Template.md')
				.setValue(this.plugin.settings.emailTemplate)
				.onChange(async (value) => {
					this.plugin.settings.emailTemplate = value;
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName('Fallback template')
			.setDesc('Path of the {{token}} template used without Templater. Empty: the built-in default.')
			.addText(text => text
				.setPlaceholder('Templates/Email Template (fallback).md')
				.setValue(this.plugin.settings.fallbackTemplate)
				.onChange(async (value) => {
					this.plugin.settings.fallbackTemplate = value;
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName('Destination folder')
			.setDesc('Folder for imported emails. Fallback only: the Templater template sets its own.')
			.addText(text => text
				.setPlaceholder(DEFAULT_SETTINGS.importFolder)
				.setValue(this.plugin.settings.importFolder)
				.onChange(async (value) => {
					this.plugin.settings.importFolder = value;
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName('File name prefix')
			.setDesc('Put before the title in the note name. Fallback only.')
			.addText(text => text
				.setPlaceholder(DEFAULT_SETTINGS.importPrefix)
				.setValue(this.plugin.settings.importPrefix)
				.onChange(async (value) => {
					this.plugin.settings.importPrefix = value;
					await this.plugin.saveSettings();
				}));

		new Setting(containerEl)
			.setName('Sent property')
			.setDesc('Property holding the sent date, used by the fallback to recognise an email that is already archived.')
			.addText(text => text
				.setPlaceholder(DEFAULT_SETTINGS.sentProperty)
				.setValue(this.plugin.settings.sentProperty)
				.onChange(async (value) => {
					this.plugin.settings.sentProperty = value;
					await this.plugin.saveSettings();
				}));
	}
}
