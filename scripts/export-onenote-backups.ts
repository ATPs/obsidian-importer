import '../tests/shims/runtime';

import * as crypto from 'node:crypto';
import * as fs from 'node:fs';
import * as path from 'node:path';
import * as url from 'node:url';

import { convertOneNoteLocal, LocalConversionReport } from './convert-onenote-local';
import { parseFrontMatterBlock, serializeFrontMatter } from '../src/util';

interface Asset {
	path: string;
	length: number;
	sha256: string;
}

interface MissingAsset {
	name: string;
	label: string;
	embed: boolean;
	record: string;
}

interface PageNote {
	file: string;
	relative: string;
	frontMatter: Record<string, unknown>;
	body: string;
	version: SourceVersion;
}

interface SourceVersion {
	notebook: string;
	source: string;
	sha256: string;
	backupMtimeMs: number;
	sectionName: string;
	sectionId?: string;
	stage: string;
	report: LocalConversionReport;
	pages: PageNote[];
	archive?: PageNote;
}

interface PageGroup {
	candidates: PageNote[];
}

interface CopiedAsset {
	asset: Asset;
	markdownPath: string;
}

interface NotebookReport {
	notebook: string;
	sources: { path: string, sha256: string, staging: string, section: string }[];
	sections: {
		name: string;
		pages: number;
		recoveredAssets: number;
		unrecoveredAssets: number;
		skippedOlderOnlyPages: number;
	}[];
	issues: string[];
}

const ONE_EXTENSIONS = new Set(['.one', '.onepkg', '.onex']);

function sha256File(file: string): string {
	return crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
}

function posix(value: string): string {
	return value.replaceAll('\\', '/');
}

function safeName(value: string): string {
	return value.normalize('NFC').replace(/[\\/:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim() || 'Untitled';
}

/** Removes OneNote's localized backup suffix without depending on its date format. */
export function sectionNameFromBackupFile(file: string): string {
	const name = path.basename(file);
	return name
		.replace(/\.one\s*\(于 [^)]*\)\.one$/iu, '')
		.replace(/\s*\(于 [^)]*\)\.one$/iu, '')
		.replace(/\.one$/iu, '')
		.trim() || 'Untitled';
}

function normalized(value: string): string {
	return value.normalize('NFC').trim().toLocaleLowerCase();
}

function pageFallbackKey(note: PageNote): string | undefined {
	const title = note.frontMatter.title;
	const created = note.frontMatter['onenote-created'];
	return typeof title === 'string' && typeof created === 'string'
		? `${normalized(title)}\u0000${created}`
		: undefined;
}

function pageTitleKey(note: PageNote): string | undefined {
	const title = note.frontMatter.title;
	return typeof title === 'string' && title !== '' ? normalized(title) : undefined;
}

function pageId(note: PageNote): string | undefined {
	const value = note.frontMatter['onenote-id'];
	return typeof value === 'string' && value !== '' ? value : undefined;
}

function pageUpdated(note: PageNote): number {
	const value = note.frontMatter['onenote-updated'];
	const parsed = typeof value === 'string' ? Date.parse(value) : Number.NaN;
	return Number.isFinite(parsed) ? parsed : note.version.backupMtimeMs;
}

function readAssets(frontMatter: Record<string, unknown>): Asset[] {
	const raw = frontMatter['onenote-assets'];
	if (!Array.isArray(raw)) return [];
	return raw.flatMap(value => {
		if (!value || typeof value !== 'object') return [];
		const candidate = value as Partial<Asset>;
		return typeof candidate.path === 'string' && typeof candidate.length === 'number' && typeof candidate.sha256 === 'string'
			? [{ path: posix(candidate.path), length: candidate.length, sha256: candidate.sha256 }]
			: [];
	});
}

function within(root: string, candidate: string): boolean {
	const relative = path.relative(root, candidate);
	return relative !== '' && !relative.startsWith('..') && !path.isAbsolute(relative);
}

function stageAsset(note: PageNote, asset: Asset): string | undefined {
	const target = path.resolve(note.version.stage, ...asset.path.split('/'));
	return within(note.version.stage, target) && fs.existsSync(target) ? target : undefined;
}

function xmlUnescape(value: string): string {
	return value.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
}

function detail(record: string, name: string): string | undefined {
	const expression = new RegExp(`<detail name="${name}" value="([\\s\\S]*?)"\\/>`);
	const match = expression.exec(record);
	return match ? xmlUnescape(match[1]) : undefined;
}

function missingAssets(body: string): MissingAsset[] {
	const result: MissingAsset[] = [];
	for (const match of body.matchAll(/<record code="ONENOTE_ASSET_DATA_MISSING">([\s\S]*?)<\/record>/g)) {
		const record = match[0];
		const name = detail(record, 'name');
		const label = detail(record, 'label');
		const embed = detail(record, 'embed');
		if (name !== undefined && label !== undefined && embed !== undefined) {
			result.push({ name, label, embed: embed === 'true', record });
		}
	}
	return result;
}

function assetBaseName(asset: Asset): string {
	const name = path.posix.basename(asset.path);
	const extension = path.posix.extname(name);
	const base = extension ? name.slice(0, -extension.length) : name;
	return base.replace(/ (\d+)$/, '') + extension;
}

function linkTargetForAsset(body: string, asset: Asset): { label: string, embed: boolean } | undefined {
	const encoded = encodeURI(asset.path);
	const at = body.indexOf(encoded);
	if (at < 0) return undefined;

	// Source folder names can contain parentheses, which are legal inside an
	// encoded Markdown destination but defeat a simplistic `...)` regex.
	const image = body.lastIndexOf('![', at);
	const link = body.lastIndexOf('[', at);
	const opening = image >= 0 && image + 1 === link ? image : link;
	if (opening < 0) return undefined;
	const embed = opening === image;
	const labelStart = opening + (embed ? 2 : 1);
	const labelEnd = body.indexOf('](', labelStart);
	if (labelEnd < labelStart || labelEnd > at) return undefined;
	return { label: body.slice(labelStart, labelEnd), embed };
}

function sourceAssetForMissing(note: PageNote, missing: MissingAsset, used: Set<string>): Asset | undefined {
	for (const asset of readAssets(note.frontMatter)) {
		const key = `${note.file}\u0000${asset.path}`;
		if (used.has(key) || assetBaseName(asset) !== missing.name) continue;
		const link = linkTargetForAsset(note.body, asset);
		if (!link || link.embed !== missing.embed || link.label !== missing.label) continue;
		if (!stageAsset(note, asset)) continue;
		used.add(key);
		return asset;
	}
	return undefined;
}

function assetDestinationName(asset: Asset): string {
	return path.basename(asset.path);
}

function withHash(name: string, hash: string, index = 0): string {
	const extension = path.extname(name);
	const base = extension ? name.slice(0, -extension.length) : name;
	return `${base}-${hash.slice(0, 8)}${index ? `-${index}` : ''}${extension}`;
}

function copyAsset(note: PageNote, asset: Asset, noteTarget: string, buildRoot: string): CopiedAsset | undefined {
	const source = stageAsset(note, asset);
	if (!source) return undefined;

	const attachmentFolder = path.join(path.dirname(noteTarget), 'attachments');
	fs.mkdirSync(attachmentFolder, { recursive: true });
	const sourceHash = sha256File(source);
	if (sourceHash !== asset.sha256 || fs.statSync(source).size !== asset.length) return undefined;

	let name = assetDestinationName(asset);
	let target = path.join(attachmentFolder, name);
	for (let index = 0; fs.existsSync(target); index++) {
		const existingHash = sha256File(target);
		if (existingHash === sourceHash) break;
		name = withHash(assetDestinationName(asset), sourceHash, index);
		target = path.join(attachmentFolder, name);
	}
	if (!fs.existsSync(target)) fs.copyFileSync(source, target);

	return {
		asset: {
			path: posix(path.relative(buildRoot, target)),
			length: asset.length,
			sha256: sourceHash,
		},
		markdownPath: posix(path.relative(path.dirname(noteTarget), target)),
	};
}

function replaceAssetPath(body: string, oldPath: string, replacement: string): string {
	const encoded = encodeURI(oldPath);
	return body.replaceAll(encoded, encodeURI(replacement)).replaceAll(oldPath, replacement);
}

function appendRecoveredLinks(body: string, links: string[]): string {
	if (links.length === 0) return body;
	return `${body.trimEnd()}\n\n## Recovered OneNote attachments\n\n${links.join('\n\n')}\n`;
}

function removeRecoveredMissingRecord(body: string, missing: MissingAsset): string {
	return body.replace(missing.record, '');
}

function targetRelative(note: PageNote): string {
	const relative = posix(path.relative(note.version.stage, note.file));
	const [section, ...rest] = relative.split('/');
	return rest.length > 0 ? rest.join('/') : path.basename(note.file);
}

function readVersion(notebook: string, source: string, stage: string, report: LocalConversionReport): SourceVersion {
	const version: SourceVersion = {
		notebook,
		source,
		sha256: sha256File(source),
		backupMtimeMs: fs.statSync(source).mtimeMs,
		sectionName: sectionNameFromBackupFile(source),
		stage,
		report,
		pages: [],
	};

	const markdown: string[] = [];
	const walk = (directory: string) => {
		for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
			const item = path.join(directory, entry.name);
			if (entry.isDirectory()) walk(item);
			else if (entry.isFile() && entry.name.toLowerCase().endsWith('.md')) markdown.push(item);
		}
	};
	walk(stage);

	for (const file of markdown) {
		const parsed = parseFrontMatterBlock(fs.readFileSync(file, 'utf8'));
		if (!parsed) continue;
		const note: PageNote = { file, relative: posix(path.relative(stage, file)), frontMatter: parsed.frontMatter, body: parsed.body, version };
		if (path.basename(file).toLowerCase() === '_onenote archive.md') version.archive = note;
		else if (pageId(note)) version.pages.push(note);
	}

	version.sectionId = typeof version.pages[0]?.frontMatter['onenote-section-id'] === 'string'
		? version.pages[0].frontMatter['onenote-section-id'] as string
		: undefined;
	return version;
}

function connectedSections(versions: SourceVersion[]): SourceVersion[][] {
	const parent = versions.map((_, index) => index);
	const find = (index: number): number => parent[index] === index ? index : (parent[index] = find(parent[index]));
	const join = (left: number, right: number): void => {
		left = find(left);
		right = find(right);
		if (left !== right) parent[right] = left;
	};

	for (let left = 0; left < versions.length; left++) {
		for (let right = left + 1; right < versions.length; right++) {
			if (normalized(versions[left].sectionName) === normalized(versions[right].sectionName)
				|| (versions[left].sectionId && versions[left].sectionId === versions[right].sectionId)) join(left, right);
		}
	}

	const groups = new Map<number, SourceVersion[]>();
	versions.forEach((version, index) => {
		const group = groups.get(find(index)) ?? [];
		group.push(version);
		groups.set(find(index), group);
	});
	return [...groups.values()];
}

function pageGroups(versions: SourceVersion[]): PageGroup[] {
	const pages = versions.flatMap(version => version.pages);
	const parent = pages.map((_, index) => index);
	const find = (index: number): number => parent[index] === index ? index : (parent[index] = find(parent[index]));
	const join = (left: number, right: number): void => {
		left = find(left);
		right = find(right);
		if (left !== right) parent[right] = left;
	};
	const ids = new Map<string, number>();
	const fallbacks = new Map<string, number>();
	const titles = new Map<string, number[]>();
	pages.forEach((page, index) => {
		const id = pageId(page);
		if (id) {
			const previous = ids.get(id);
			if (previous === undefined) ids.set(id, index);
			else join(index, previous);
		}
		const fallback = pageFallbackKey(page);
		if (fallback) {
			const previous = fallbacks.get(fallback);
			if (previous === undefined) fallbacks.set(fallback, index);
			else join(index, previous);
		}
		const title = pageTitleKey(page);
		if (title) {
			const matches = titles.get(title) ?? [];
			matches.push(index);
			titles.set(title, matches);
		}
	});

	// Some backup rewrites replace page IDs and omit the creation timestamp. A
	// title is only a valid fallback when it appears once in every candidate
	// backup; duplicate titles intentionally remain separate pages.
	for (const indexes of titles.values()) {
		const byVersion = new Map<SourceVersion, number[]>();
		for (const index of indexes) {
			const group = byVersion.get(pages[index].version) ?? [];
			group.push(index);
			byVersion.set(pages[index].version, group);
		}
		if ([...byVersion.values()].some(group => group.length !== 1)) continue;
		const unique = [...byVersion.values()].map(group => group[0]);
		for (let index = 1; index < unique.length; index++) join(unique[0], unique[index]);
	}
	const groups = new Map<number, PageNote[]>();
	pages.forEach((page, index) => {
		const group = groups.get(find(index)) ?? [];
		group.push(page);
		groups.set(find(index), group);
	});
	return [...groups.values()].map(candidates => ({ candidates }));
}

function selectedPage(group: PageGroup): PageNote {
	return [...group.candidates].sort((left, right) =>
		pageUpdated(right) - pageUpdated(left)
		|| right.version.backupMtimeMs - left.version.backupMtimeMs
		|| right.file.localeCompare(left.file))[0];
}

function markdownLink(missing: MissingAsset, target: string): string {
	return missing.embed ? `![${missing.label}](${encodeURI(target)})` : `[${missing.label}](${encodeURI(target)})`;
}

function writePage(
	selected: PageNote,
	candidates: PageNote[],
	target: string,
	buildRoot: string,
	issues: string[],
): { recovered: number, unrecovered: number } {
	let body = selected.body;
	const frontMatter = { ...selected.frontMatter };
	const copied: Asset[] = [];
	for (const asset of readAssets(frontMatter)) {
		const result = copyAsset(selected, asset, target, buildRoot);
		if (!result) {
			issues.push(`${selected.file}: current asset unavailable or failed integrity check: ${asset.path}`);
			continue;
		}
		copied.push(result.asset);
		body = replaceAssetPath(body, asset.path, result.markdownPath);
	}

	const used = new Set<string>();
	const recoveredLinks: string[] = [];
	let recovered = 0;
	let unrecovered = 0;
	for (const missing of missingAssets(body)) {
		const older = candidates
			.filter(candidate => candidate !== selected)
			.sort((left, right) => right.version.backupMtimeMs - left.version.backupMtimeMs);
		let result: CopiedAsset | undefined;
		for (const candidate of older) {
			const source = sourceAssetForMissing(candidate, missing, used);
			if (!source) continue;
			result = copyAsset(candidate, source, target, buildRoot);
			if (result) break;
		}
		if (!result) {
			unrecovered++;
			issues.push(`${selected.file}: could not recover missing ${missing.name}`);
			continue;
		}
		copied.push(result.asset);
		recoveredLinks.push(markdownLink(missing, result.markdownPath));
		body = removeRecoveredMissingRecord(body, missing);
		recovered++;
	}

	frontMatter['onenote-assets'] = copied;
	frontMatter['onenote-completeness'] = unrecovered > 0 ? 'degraded' : recovered > 0 ? 'recovered' : frontMatter['onenote-completeness'];
	frontMatter['onenote-merged-from'] = selected.version.source;
	frontMatter['onenote-recovered-assets'] = recovered;
	body = appendRecoveredLinks(body, recoveredLinks);
	fs.mkdirSync(path.dirname(target), { recursive: true });
	fs.writeFileSync(target, serializeFrontMatter(frontMatter) + body, 'utf8');
	return { recovered, unrecovered };
}

function writeArchive(source: PageNote, target: string, buildRoot: string, issues: string[]): void {
	const frontMatter = { ...source.frontMatter };
	let body = source.body;
	const copied: Asset[] = [];
	for (const asset of readAssets(frontMatter)) {
		const result = copyAsset(source, asset, target, buildRoot);
		if (!result) {
			issues.push(`${source.file}: archive asset unavailable or failed integrity check: ${asset.path}`);
			continue;
		}
		copied.push(result.asset);
		body = replaceAssetPath(body, asset.path, result.markdownPath);
	}
	frontMatter['onenote-assets'] = copied;
	frontMatter['onenote-merged-from'] = source.version.source;
	fs.mkdirSync(path.dirname(target), { recursive: true });
	fs.writeFileSync(target, serializeFrontMatter(frontMatter) + body, 'utf8');
}

function auditBuild(root: string): string[] {
	const errors: string[] = [];
	const markdown: string[] = [];
	const walk = (directory: string) => {
		for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
			const item = path.join(directory, entry.name);
			if (entry.isDirectory()) walk(item);
			else if (entry.isFile() && entry.name.toLowerCase().endsWith('.md')) markdown.push(item);
		}
	};
	walk(root);
	for (const file of markdown) {
		const parsed = parseFrontMatterBlock(fs.readFileSync(file, 'utf8'));
		if (!parsed) continue;
		for (const asset of readAssets(parsed.frontMatter)) {
			const assetPath = path.resolve(root, ...asset.path.split('/'));
			if (!within(root, assetPath) || !fs.existsSync(assetPath)) {
				errors.push(`${file}: missing asset ${asset.path}`);
				continue;
			}
			if (fs.statSync(assetPath).size !== asset.length || sha256File(assetPath) !== asset.sha256) {
				errors.push(`${file}: asset integrity mismatch ${asset.path}`);
			}
		}
	}
	return errors;
}

function inputFiles(directory: string): string[] {
	return fs.readdirSync(directory, { withFileTypes: true })
		.filter(entry => entry.isFile() && ONE_EXTENSIONS.has(path.extname(entry.name).toLowerCase()))
		.map(entry => path.join(directory, entry.name));
}

/** Export backup versions separately, then publish one latest-text merged notebook per input directory. */
export async function exportOneNoteBackups(outputArgument: string, inputArguments: string[], reuseStaging = false): Promise<NotebookReport[]> {
	const outputRoot = path.resolve(outputArgument);
	const inputs = inputArguments.map(input => path.resolve(input));
	if (inputs.length === 0) throw new Error('Give at least one OneNote backup directory.');
	for (const input of inputs) {
		if (!fs.statSync(input).isDirectory()) throw new Error(`Not a directory: ${input}`);
	}
	fs.mkdirSync(outputRoot, { recursive: true });

	const notebookNames = inputs.map(input => path.basename(input));
	if (new Set(notebookNames.map(normalized)).size !== notebookNames.length) throw new Error('Input directories must have distinct names.');
	for (const notebook of notebookNames) {
		const target = path.join(outputRoot, notebook);
		if (fs.existsSync(target) && !reuseStaging) throw new Error(`Final output already exists: ${target}`);
	}

	const allVersions: SourceVersion[] = [];
	for (const input of inputs) {
		const notebook = path.basename(input);
		const files = inputFiles(input);
		if (files.length === 0) throw new Error(`No OneNote files found in ${input}`);
		for (const [index, source] of files.entries()) {
			const hash = sha256File(source);
			const stage = path.join(outputRoot, '_staging', notebook, `${String(index + 1).padStart(3, '0')}-${safeName(path.basename(source))}-${hash.slice(0, 8)}`);
			let report: LocalConversionReport;
			if (fs.existsSync(stage)) {
				if (!reuseStaging) throw new Error(`Staging output already exists: ${stage}`);
				const reportFile = path.join(stage, '_conversion-report.json');
				if (!fs.existsSync(reportFile)) throw new Error(`Staging report is missing: ${reportFile}`);
				report = JSON.parse(fs.readFileSync(reportFile, 'utf8')) as LocalConversionReport;
			}
			else {
				fs.mkdirSync(path.dirname(stage), { recursive: true });
				report = await convertOneNoteLocal(source, stage);
			}
			allVersions.push(readVersion(notebook, source, stage, report));
		}
	}

	const reports: NotebookReport[] = [];
	for (const notebook of notebookNames) {
		const versions = allVersions.filter(version => version.notebook === notebook);
		const build = path.join(outputRoot, `.build-${notebook}-${Date.now()}`);
		const report: NotebookReport = {
			notebook,
			sources: versions.map(version => ({ path: version.source, sha256: version.sha256, staging: version.stage, section: version.sectionName })),
			sections: [],
			issues: [],
		};
		fs.mkdirSync(build, { recursive: false });

		for (const section of connectedSections(versions)) {
			const newest = [...section].sort((left, right) => right.backupMtimeMs - left.backupMtimeMs || right.source.localeCompare(left.source))[0];
			const sectionName = safeName(newest.sectionName);
			const sectionRoot = path.join(build, sectionName);
			let recoveredAssets = 0;
			let unrecoveredAssets = 0;
			let pages = 0;
			let skippedOlderOnlyPages = 0;
			for (const group of pageGroups(section)) {
				// A page absent from the newest section backup was deleted. Its older
				// markdown and attachments stay in staging but never reach final output.
				if (!group.candidates.some(candidate => candidate.version === newest)) {
					skippedOlderOnlyPages++;
					continue;
				}
				const selected = selectedPage(group);
				const target = path.join(sectionRoot, ...targetRelative(selected).split('/'));
				const outcome = writePage(selected, group.candidates, target, build, report.issues);
				recoveredAssets += outcome.recovered;
				unrecoveredAssets += outcome.unrecovered;
				pages++;
			}
			const archive = [...section].map(version => version.archive).filter((note): note is PageNote => note !== undefined)
				.sort((left, right) => right.version.backupMtimeMs - left.version.backupMtimeMs)[0];
			if (archive) writeArchive(archive, path.join(sectionRoot, '_OneNote archive.md'), build, report.issues);
			report.sections.push({ name: sectionName, pages, recoveredAssets, unrecoveredAssets, skippedOlderOnlyPages });
		}

		const auditErrors = auditBuild(build);
		if (auditErrors.length > 0) throw new Error(`Refusing to publish ${notebook}; ${auditErrors.join('\n')}`);
		fs.writeFileSync(path.join(build, '_merge-report.json'), JSON.stringify(report, null, 2), 'utf8');
		const final = path.join(outputRoot, notebook);
		if (fs.existsSync(final)) {
			const backup = path.join(outputRoot, `.superseded-${notebook}-${Date.now()}`);
			fs.renameSync(final, backup);
		}
		fs.renameSync(build, final);
		reports.push(report);
	}

	fs.writeFileSync(path.join(outputRoot, '_merge-report.json'), JSON.stringify(reports, null, 2), 'utf8');
	return reports;
}

async function main(): Promise<void> {
	const arguments_ = process.argv.slice(2);
	const reuseStaging = arguments_.includes('--reuse-staging');
	const positional = arguments_.filter(argument => argument !== '--reuse-staging');
	const [output, ...inputs] = positional;
	if (!output || inputs.length === 0) {
		throw new Error('Usage: export-onenote-backups.ts <output directory> [--reuse-staging] <backup directory> [...]');
	}
	const reports = await exportOneNoteBackups(output, inputs, reuseStaging);
	process.stdout.write(JSON.stringify(reports.map(report => ({
		notebook: report.notebook,
		sections: report.sections.length,
		recoveredAssets: report.sections.reduce((total, section) => total + section.recoveredAssets, 0),
	})), null, 2) + '\n');
}

if (process.argv[1] && path.resolve(process.argv[1]) === path.resolve(url.fileURLToPath(import.meta.url))) {
	void main().catch(error => {
		console.error(error);
		process.exitCode = 1;
	});
}
