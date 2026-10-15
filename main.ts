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
 */
import { App, Plugin, PluginSettingTab, Setting, Notice, MarkdownRenderer, TFile, Component, normalizePath, FileSystemAdapter, Editor, MarkdownView, MarkdownFileInfo, getFrontMatterInfo, getLinkpath, resolveSubpath } from 'obsidian';
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
}

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
	safetyPlaceholder: 'REMOVE-BEFORE-SENDING'
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

export default class MailForOutlookPlugin extends Plugin {
	settings: MailForOutlookSettings;
	// True while an export runs, so a second one can't start
	private exporting = false;

	async onload() {
		await this.loadSettings();

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

		this.addSettingTab(new MailForOutlookSettingTab(this.app, this));
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
					// Show or hide the person note settings
					this.display();
				}));

		if (this.plugin.settings.recipientsFromPersonNotes) {
			new Setting(containerEl)
				.setName('Person note type')
				.setDesc('Linked notes whose "type" property has this value are treated as people.')
				.addText(text => text
					.setPlaceholder(DEFAULT_SETTINGS.personType)
					.setValue(this.plugin.settings.personType)
					.onChange(async (value) => {
						this.plugin.settings.personType = value;
						await this.plugin.saveSettings();
					}));

			new Setting(containerEl)
				.setName('Email property')
				.setDesc('Property in person notes that holds their email address.')
				.addText(text => text
					.setPlaceholder(DEFAULT_SETTINGS.emailProperty)
					.setValue(this.plugin.settings.emailProperty)
					.onChange(async (value) => {
						this.plugin.settings.emailProperty = value;
						await this.plugin.saveSettings();
					}));
		}

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
	}
}
