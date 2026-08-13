import { Notice, Platform, TFile, TFolder, normalizePath } from 'obsidian';
import { helpUrl } from '../constants';
import { PickedFile, fs, os, path as nodePath } from '../filesystem';
import { FormatImporter, PlannedNote, leavesTheNoteAlone } from '../format-importer';
import { ImportContext } from '../import-context';
import { i18n } from '../i18n';
import { selectedNodes } from '../tree';
import { TreePicker, ViewableNode } from '../tree-view';
import { describeReason, extensionFromBytes, parseFrontMatterBlock, sanitizeFileName, serializeFrontMatter, uint8arrayToArrayBuffer } from '../util';
import { findBackupFolder } from './onenote-file/backup-folder';
import { AssetSource, convertPage } from './onenote-file/convert';
import { OneNoteErrorKind, OneNoteFormatError } from './onenote-file/errors';
import { isPackage, listSections, readSections } from './onenote-file/package';
import { sha256Hex } from './onenote-file/preservation';
import { Page, Section } from './onenote-file/semantic/content';
import { retryTransient } from './onenote-file/retry';

const HELP_PERMALINK = 'import/onenote';

interface SectionNode extends ViewableNode<SectionNode> {
	file: PickedFile;
	entryName?: string;
}

interface SourceIdentity {
	path: string;
	sha256: string;
}

interface PreparedSection {
	section: Section;
	title: string;
	groups: string[];
	source: SourceIdentity;
}

interface AssetManifestEntry {
	path: string;
	length: number;
	sha256: string;
	sourceName?: string;
	ordinal?: number;
	embed?: boolean;
}

interface GeneratedAsset extends AssetManifestEntry {
	created: boolean;
}

interface PlannedSection {
	item: PreparedSection;
	folderPath: string;
	pages: Map<Page, PlannedPage>;
}

interface PlannedPage {
	note: PlannedNote;
	/** Readable sanitized name before an actual Windows path limit truncates it. */
	fullOutputFilename: string;
}

const REASONS: Record<OneNoteErrorKind, () => string> = {
	unsupported: () => i18n.importer.onenoteFile.reasonUnsupported(),
	protected: () => i18n.importer.onenoteFile.reasonRightsProtected(),
	malformed: () => i18n.importer.onenoteFile.reasonMalformed(),
	limit: () => i18n.importer.onenoteFile.reasonTooLarge(),
};

export class OneNoteFileImporter extends FormatImporter {
	static extensions = ['one', 'onepkg', 'onex'];

	interruption = 'pause' as const;

	// Field initializers would overwrite values set by base-constructor init().
	private picker: TreePicker<SectionNode>;
	private loadedFrom = '';
	private loadGeneration = 0;

	init(): void {
		this.addSetting('source')
			?.setName(i18n.common.nameExport())
			.setDesc(i18n.importer.onenoteFile.descExport())
			.addButton(button => button
				.setButtonText(i18n.common.buttonInstructions())
				.onClick(() => window.open(helpUrl(HELP_PERMALINK))));

		const backup = windowsBackupFolder();
		this.addFileChooserSetting(
			i18n.importer.onenoteFile.fileType(),
			OneNoteFileImporter.extensions,
			true,
			backup ? i18n.importer.onenoteFile.descBackupFolder() : undefined,
			backup);
		this.defaultOutputFolder = 'OneNote';
		this.idProperty = 'onenote-id';
		this.idLabel = i18n.importer.onenoteFile.labelId();

		this.drawSectionPicker();
	}

	protected sourceChanged(): void {
		super.sourceChanged();

		const key = this.files.map(file => file.fullpath).join('\n');
		if (key === this.loadedFrom) return;

		this.loadedFrom = key;
		if (this.picker) void this.loadSections();
	}

	private drawSectionPicker(): void {
		this.draw(contentEl => {
			this.picker = new TreePicker<SectionNode>(contentEl, {
				name: i18n.importer.onenoteFile.nameSections(),
				desc: i18n.importer.onenoteFile.descSections(),
				hint: i18n.importer.onenoteFile.msgPickFileFirst(),
				loading: i18n.importer.onenoteFile.msgLoadingSections(),
				empty: i18n.importer.onenoteFile.msgNoSections(),
				failed: describeFailure,
				view: {
					icon: node => node.children?.length ? 'book' : 'file-text',
				},
				loadsItself: true,
			});
		}, 'source');
	}

	private async loadSections(): Promise<void> {
		if (this.files.length === 0) {
			this.picker.reset();
			return;
		}

		const generation = ++this.loadGeneration;

		await this.picker.load(async () => {
			const nodes: SectionNode[] = [];

			for (const file of this.files) {
				const data = new Uint8Array(await retryTransient(() => file.read()));

				// Ignore a read superseded while it was in progress.
				if (generation !== this.loadGeneration) return this.picker.nodes;

				const sections = listSections(data, file.name);

				if (!isPackage(data)) {
					nodes.push({ title: file.basename, file, selected: true, disabled: false });
					continue;
				}

				nodes.push({
					title: file.basename,
					file,
					selected: true,
					disabled: false,
					children: sections.map(entry => ({
						title: entry.title,
						file,
						entryName: entry.name,
						selected: true,
						disabled: false,
					})),
				});
			}

			return generation === this.loadGeneration ? nodes : this.picker.nodes;
		});
	}

	async import(ctx: ImportContext): Promise<void> {
		if (this.files.length === 0) {
			new Notice(i18n.common.msgPickFile());
			return;
		}

		const folder = await this.getOutputFolder();
		if (!folder) {
			new Notice(i18n.common.msgPickOutput());
			return;
		}

		const nodes = this.picker?.nodes ?? [];

		// A missing picker means all sections; an empty selection means none.
		const loaded = nodes.length > 0;
		const chosen = selectedNodes(nodes, node => !node.children?.length);

		const prepared: PreparedSection[] = [];
		for (const file of this.files) {
			if (await ctx.shouldStop()) return;

			const forThisFile = chosen.filter(node => node.file === file);
			if (loaded && forThisFile.length === 0) continue;

			const wanted = loaded
				? new Set(forThisFile.map(node => node.entryName).filter((name): name is string => name !== undefined))
				: undefined;

			try {
				prepared.push(...await this.prepareFile(ctx, file, wanted));
			}
			catch (error) {
				report(ctx, `${file.name} (${file.toString()})`, error);
			}
		}

		const planned = this.planSections(prepared, folder);
		const links = oneNoteLinkNames(
			prepared.flatMap(item => item.section.pages),
			new Map([...planned].flatMap(section => [...section.pages].map(([page, plan]) => [page.id, linkTarget(plan.note.targetPath)]))));
		for (const section of planned) {
			if (await ctx.shouldStop()) return;
			try {
				await this.importSection(ctx, section, links);
			}
			catch (error) {
				const { item } = section;
				report(ctx, `${item.title} (${item.source.path}; SHA-256 ${item.source.sha256})`, error);
			}
		}
	}

	private planSections(prepared: PreparedSection[], root: TFolder): PlannedSection[] {
		const names = oneNotePageNames(prepared.flatMap(item => item.section.pages));
		const planned: PlannedSection[] = [];

		for (const item of prepared) {
			let folderPath = root.path === '/' ? '' : root.path;
			for (const group of item.groups) folderPath = appendName(folderPath, sanitizeFileName(group, folderPath));
			folderPath = appendName(folderPath, sanitizeFileName(item.section.name || item.title, folderPath));

			const pages = new Map<Page, PlannedPage>();
			const levels: string[] = [folderPath];
			for (const page of item.section.pages) {
				const depth = Math.min(page.level, levels.length - 1);
				levels.length = depth + 1;

				let parent = levels[depth];
				if (page.isDeleted) parent = appendName(folderPath, '_deleted');
				else if (page.isConflictPage) parent = appendName(folderPath, '_conflicts');

				const proposed = names.get(page.id)!;
				const actual = sanitizeFileName(proposed, parent);
				const timestamp = oneNoteTimestamp(page);
				const title = actual === proposed
					? proposed
					: sanitizeFileName(actual.startsWith(`${timestamp} `) ? actual : `${timestamp} ${actual}`, parent);
				const note = this.planNote(parent, title, page.id);
				pages.set(page, { note, fullOutputFilename: `${proposed}.md` });
				levels.push(linkTarget(note.targetPath));
			}
			planned.push({ item, folderPath, pages });
		}

		return planned;
	}

	private async prepareFile(ctx: ImportContext, file: PickedFile, wanted?: Set<string>): Promise<PreparedSection[]> {
		ctx.status(i18n.importer.onenoteFile.statusReadingSection({ name: file.name }));

		const data = new Uint8Array(await retryTransient(() => file.read(), {
			onRetry: attempt => ctx.reportMessage(`Retrying OneNote source read for ${file.name} (attempt ${attempt}/3).`),
		}));
		const source: SourceIdentity = { path: file.toString(), sha256: await sha256Hex(data) };
		let sourceModifiedUtc: Date | undefined;
		try {
			sourceModifiedUtc = new Date(fs.statSync(source.path).mtimeMs);
		}
		catch {
			// Browser-picked and packaged files do not necessarily expose a local path.
		}
		let sections;
		try {
			sections = readSections(data, file.name, wanted?.size ? wanted : undefined);
		}
		catch (error) {
			report(ctx, `${file.name} (${source.path}; ${data.length} bytes; SHA-256 ${source.sha256})`, error);
			return [];
		}
		const prepared: PreparedSection[] = [];
		let done = 0;

		for (const entry of sections) {
			if (await ctx.shouldStop()) return prepared;

			ctx.status(i18n.importer.onenoteFile.statusImportingSection({
				name: entry.title,
				index: ++done,
				total: sections.length,
			}));

			let section: Section;
			try {
				section = entry.read();
			}
			catch (error) {
				report(ctx, `${entry.title} (${source.path}; ${data.length} bytes; SHA-256 ${source.sha256})`, error);
				continue;
			}
			if (sourceModifiedUtc) {
				for (const page of section.pages) {
					if (!page.createdUtc && !page.lastModifiedUtc) page.lastModifiedUtc = sourceModifiedUtc;
				}
			}

			prepared.push({ section, title: entry.title, groups: entry.groups, source });
		}

		return prepared;
	}

	private async importSection(
		ctx: ImportContext,
		planned: PlannedSection,
		links: Map<string, string[]>,
	): Promise<void> {
		const { item, pages } = planned;
		const { section, source } = item;
		let done = 0;

		for (const page of section.pages) {
			if (await ctx.shouldStop()) return;

			ctx.reportProgress(++done, section.pages.length);
			const plan = pages.get(page)!;
			const { note } = plan;
			await this.createFolders(note.targetPath.slice(0, note.targetPath.lastIndexOf('/')) || '/');
			await this.importPage(ctx, page, plan, links, section.id, source);
		}
	}

	private async importPage(
		ctx: ImportContext,
		page: Page,
		plan: PlannedPage,
		links: Map<string, string[]>,
		sectionId: string,
		source: SourceIdentity,
	): Promise<string | undefined> {
		const { note: planned } = plan;
		const title = planned.title;
		const generatedAssets: GeneratedAsset[] = [];

		try {
			// Preflight before conversion so skipped notes write no attachments.
			const disposition = this.preflightNote(ctx, planned, page.lastModifiedUtc?.getTime());
			if (leavesTheNoteAlone(disposition)) return title;

			const notePath = planned.targetPath;
			const oldAssets = planned.file ? assetsIn(await this.vault.read(planned.file)) : [];
			const save = async (bytes: Uint8Array, suggested: string, source?: AssetSource) => {
				const saved = await this.saveAttachment(ctx, bytes, suggested, notePath);
				const asset = { path: saved.path, length: bytes.length, sha256: await sha256Hex(bytes), ...source };
				generatedAssets.push({ ...asset, created: saved.created });
				return { name: saved.name, ...asset };
			};
			const converted = await convertPage(page, {
				noteName: title,
				isCancelled: () => ctx.isCancelled(),
				resolveInternalLink: pageTitle => links.get(pageTitle.normalize('NFC').toLocaleLowerCase()),
				onSkipped: (name, reason) => ctx.reportSkipped(name, reason === 'no-data'
					? i18n.importer.onenoteFile.reasonNoAttachmentData()
					: i18n.importer.onenoteFile.reasonNotRepresentable()),
				saveAttachment: save,
			});
			if (converted.cancelled) {
				await this.removeUnownedGeneratedAssets(ctx, generatedAssets);
				return undefined;
			}
			const frontMatter = serializeFrontMatter({
				title: page.title,
				'onenote-original-filename': page.title === '' ? '' : `${page.title}.md`,
				'onenote-full-output-filename': plan.fullOutputFilename,
				'onenote-output-filename': planned.targetPath.slice(planned.targetPath.lastIndexOf('/') + 1),
				'onenote-id': page.id,
				'onenote-section-id': sectionId,
				'onenote-source': source.path,
				'onenote-source-sha256': source.sha256,
				'onenote-status': page.isDeleted ? 'deleted' : page.isConflictPage ? 'conflict' : 'current',
				'onenote-level': page.level,
				'onenote-created': page.createdUtc?.toISOString(),
				'onenote-updated': page.lastModifiedUtc?.toISOString(),
				'onenote-completeness': converted.degraded ? 'degraded' : 'complete',
				'onenote-assets': assetManifest(generatedAssets),
				'onenote-missing-assets': converted.missingAssets,
				'onenote-html-fallbacks': converted.htmlFallbacks || undefined,
			});

			const { written } = await this.writePlannedWithRetry(
				ctx,
				planned,
				frontMatter + converted.markdown,
				{
					sourceId: page.id,
					ctime: page.createdUtc?.getTime(),
					mtime: page.lastModifiedUtc?.getTime(),
					disposition,
				},
				`OneNote note write for ${title}`);

			if (written) {
				this.reportGeneratedAssets(ctx, generatedAssets);
				ctx.reportNoteSuccess(title);
				try {
					await this.removeStaleAssets(ctx, oldAssets, generatedAssets, planned.targetPath);
				}
				catch (error) {
					ctx.reportFailed(`${title}: stale OneNote attachment cleanup`, error);
				}
			}
			return title;
		}
		catch (error) {
			await this.removeUnownedGeneratedAssets(ctx, generatedAssets);
			ctx.reportFailed(title, error);
			return undefined;
		}
	}

	private async writePlannedWithRetry(
		ctx: ImportContext,
		planned: ReturnType<OneNoteFileImporter['planNote']>,
		content: string,
		options: NonNullable<Parameters<OneNoteFileImporter['writePlannedNote']>[3]>,
		label: string,
	) {
		let retried = false;
		return await retryTransient(async () => {
			if (retried && !planned.file && options.sourceId && this.idProperty) {
				const appeared = this.vault.getAbstractFileByPath(planned.targetPath);
				if (appeared instanceof TFile) {
					const recorded = this.sourceIdIn(await this.vault.read(appeared), this.idProperty);
					if (recorded === options.sourceId) {
						this.trackMarkdownFile(appeared);
						return { file: appeared, written: true, outcome: 'created' as const };
					}
				}
			}

			return await this.writePlannedNote(ctx, planned, content, options);
		}, {
			onRetry: attempt => {
				retried = true;
				ctx.reportMessage(`Retrying ${label} (attempt ${attempt}/3).`);
			},
		});
	}

	private async saveAttachment(ctx: ImportContext, bytes: Uint8Array, suggested: string, notePath: string) {
		const data = uint8arrayToArrayBuffer(bytes as Uint8Array<ArrayBuffer>);

		if (!/\.[^.\\/]+$/.test(suggested)) {
			const sniffed = extensionFromBytes(bytes);
			if (sniffed) suggested = `${suggested}.${sniffed}`;
		}

		const { path, reuse } = await this.placeAttachment(suggested, notePath, async (existing: TFile) => {
			if (existing.stat.size !== data.byteLength) return 'another';
			const onDisk = new Uint8Array(await this.vault.readBinary(existing));
			return onDisk.every((byte, index) => byte === bytes[index]) ? 'same' : 'another';
		});

		const created = !reuse;
		if (created) {
			await retryTransient(() => this.writeAttachment(path, data), {
				onRetry: attempt => ctx.reportMessage(`Retrying OneNote attachment write for ${suggested} (attempt ${attempt}/3).`),
			});
		}

		return { path: reuse?.path ?? path, name: suggested, created };
	}

	private reportGeneratedAssets(ctx: ImportContext, assets: GeneratedAsset[]): void {
		for (const asset of uniqueAssets(assets)) {
			if (asset.created) ctx.reportAttachmentSuccess(asset.path);
		}
	}

	private async removeUnownedGeneratedAssets(ctx: ImportContext, assets: GeneratedAsset[]): Promise<void> {
		const owned = new Set<string>();
		try {
			for (const note of this.vault.getMarkdownFiles()) {
				for (const asset of assetsIn(await this.vault.read(note))) owned.add(normalizePath(asset.path).toLocaleLowerCase());
			}
		}
		catch (error) {
			ctx.reportFailed('Unowned OneNote attachment ownership check', error);
			return;
		}

		for (const asset of uniqueAssets(assets)) {
			if (!asset.created || owned.has(normalizePath(asset.path).toLocaleLowerCase())) continue;
			try {
				const file = this.vault.getAbstractFileByPath(asset.path);
				if (!(file instanceof TFile)) continue;
				const bytes = new Uint8Array(await this.vault.readBinary(file));
				if (bytes.length !== asset.length || await sha256Hex(bytes) !== asset.sha256) {
					ctx.reportMessage(`${asset.path}: kept because it changed before failed OneNote import cleanup.`);
					continue;
				}
				await retryTransient(() => this.app.fileManager.trashFile(file), {
					onRetry: attempt => ctx.reportMessage(`Retrying unowned OneNote attachment removal for ${asset.path} (attempt ${attempt}/3).`),
				});
				ctx.reportMessage(`${asset.path}: removed because its OneNote note was not committed.`);
			}
			catch (error) {
				ctx.reportFailed(`${asset.path}: unowned OneNote attachment cleanup`, error);
			}
		}
	}

	private async removeStaleAssets(
		ctx: ImportContext,
		oldAssets: AssetManifestEntry[],
		currentAssets: GeneratedAsset[],
		ownerPath: string,
	): Promise<void> {
		const current = new Set(currentAssets.map(asset => normalizePath(asset.path).toLocaleLowerCase()));
		const shared = new Set<string>();
		for (const note of this.vault.getMarkdownFiles()) {
			if (note.path === ownerPath) continue;
			for (const asset of assetsIn(await this.vault.read(note))) shared.add(normalizePath(asset.path).toLocaleLowerCase());
		}

		for (const asset of oldAssets) {
			const key = normalizePath(asset.path).toLocaleLowerCase();
			if (current.has(key) || shared.has(key)) continue;
			const file = this.vault.getAbstractFileByPath(asset.path);
			if (!(file instanceof TFile)) continue;
			const bytes = new Uint8Array(await this.vault.readBinary(file));
			if (bytes.length !== asset.length || await sha256Hex(bytes) !== asset.sha256) {
				ctx.reportMessage(`${asset.path}: kept because it was modified after OneNote imported it.`);
				continue;
			}
			await retryTransient(() => this.app.fileManager.trashFile(file), {
				onRetry: attempt => ctx.reportMessage(`Retrying stale OneNote attachment removal for ${asset.path} (attempt ${attempt}/3).`),
			});
			ctx.reportMessage(`${asset.path}: removed because the updated OneNote page no longer references it.`);
		}
	}
}

function oneNoteTimestamp(page: Page): string {
	const date = page.createdUtc ?? page.lastModifiedUtc;
	if (!date) return 'unknown-date';
	const pad = (value: number): string => String(value).padStart(2, '0');
	return `${date.getUTCFullYear()}${pad(date.getUTCMonth() + 1)}${pad(date.getUTCDate())}-${pad(date.getUTCHours())}-${pad(date.getUTCMinutes())}-${pad(date.getUTCSeconds())}`;
}

export function oneNotePageNames(pages: Page[]): Map<string, string> {
	const base = pages.map(page => sanitizeFileName(page.title));
	const counts = new Map<string, number>();
	for (const name of base) counts.set(name.toLocaleLowerCase(), (counts.get(name.toLocaleLowerCase()) ?? 0) + 1);
	const used = new Map<string, number>();

	return new Map(pages.map((page, index) => {
		const cleaned = base[index];
		const lossy = page.title.normalize('NFC').trim() !== cleaned || (counts.get(cleaned.toLocaleLowerCase()) ?? 0) > 1;
		if (!lossy) return [page.id, cleaned];
		const baseName = `${oneNoteTimestamp(page)} ${cleaned}`;
		const key = baseName.toLocaleLowerCase();
		const duplicate = used.get(key) ?? 0;
		used.set(key, duplicate + 1);
		return [page.id, duplicate === 0 ? baseName : `${baseName}-${duplicate}`];
	}));
}

export function oneNoteLinkNames(pages: Page[], names: Map<string, string>): Map<string, string[]> {
	const found = new Map<string, string[]>();
	for (const page of pages) {
		const key = page.title.normalize('NFC').toLocaleLowerCase();
		const targets = found.get(key) ?? [];
		targets.push(names.get(page.id)!);
		found.set(key, targets);
	}
	return found;
}

function uniqueAssets(assets: GeneratedAsset[]): GeneratedAsset[] {
	return [...new Map(assets.map(asset => [normalizePath(asset.path).toLocaleLowerCase(), asset])).values()];
}

function assetManifest(assets: GeneratedAsset[]): AssetManifestEntry[] {
	return uniqueAssets(assets).map(({ path, length, sha256, sourceName, ordinal, embed }) => ({
		path, length, sha256, sourceName, ordinal, embed,
	}));
}

function assetsIn(markdown: string): AssetManifestEntry[] {
	const value: unknown = parseFrontMatterBlock(markdown)?.frontMatter['onenote-assets'];
	if (!Array.isArray(value)) return [];
	return value.flatMap(item => {
		if (!item || typeof item !== 'object') return [];
		const asset = item as Partial<AssetManifestEntry>;
		return typeof asset.path === 'string' && typeof asset.length === 'number' && typeof asset.sha256 === 'string'
			? [{
				path: asset.path,
				length: asset.length,
				sha256: asset.sha256,
				sourceName: typeof asset.sourceName === 'string' ? asset.sourceName : undefined,
				ordinal: typeof asset.ordinal === 'number' ? asset.ordinal : undefined,
				embed: typeof asset.embed === 'boolean' ? asset.embed : undefined,
			}]
			: [];
	});
}

function appendName(parent: string, name: string): string {
	return normalizePath(parent ? `${parent}/${name}` : name);
}

function linkTarget(markdownPath: string): string {
	return markdownPath.replace(/\.md$/i, '');
}


function windowsBackupFolder(): string | undefined {
	if (!Platform.isWin || !Platform.isDesktopApp || !fs || !nodePath || !os) return undefined;

	return findBackupFolder({
		root: nodePath.join(os.homedir(), 'AppData', 'Local', 'Microsoft', 'OneNote'),
		join: (...parts: string[]) => nodePath.join(...parts),
		list: directory => {
			try {
				return fs.readdirSync(directory);
			}
			catch {
				return undefined;
			}
		},
	});
}

function describeFailure(error: unknown): string {
	return error instanceof OneNoteFormatError ? REASONS[error.kind]() : describeReason(error);
}

function report(ctx: ImportContext, name: string, error: unknown): void {
	const expected = error instanceof OneNoteFormatError && (error.kind === 'protected' || error.kind === 'unsupported');

	if (expected) ctx.reportSkipped(name, describeFailure(error));
	else ctx.reportFailed(name, describeFailure(error));
}
