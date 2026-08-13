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
	/** Stable source-image name, retained for older-backup recovery. */
	sourceName?: string;
	/** Zero-based occurrence of this source name within its page. */
	ordinal?: number;
	/** Whether the source rendered as an image rather than a file link. */
	embed?: boolean;
}

interface MissingAsset {
	name: string;
	label: string;
	embed: boolean;
	/** Present on new conversions; never derived from a Markdown label. */
	sourceName?: string;
	/** Present with sourceName to distinguish same-name attachments. */
	ordinal?: number;
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
}

interface PageGroup {
	candidates: PageNote[];
}

interface CopiedAsset {
	asset: Asset;
	markdownPath: string;
}

interface PlannedPage {
	selected: PageNote;
	candidates: PageNote[];
	target: string;
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

function titleKeys(value: string): string[] {
	const title = normalized(value).replace(/\s+/gu, ' ');
	return [`title:${title}`];
}

function titleMatches(value: string, targets: Map<string, string[]>): Set<string> {
	const keys = titleKeys(value);
	const exact = targets.get(keys[0]) ?? [];
	return new Set(exact.length > 0 ? exact : keys.slice(1).flatMap(key => targets.get(key) ?? []));
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
			? [{
				path: posix(candidate.path),
				length: candidate.length,
				sha256: candidate.sha256,
				sourceName: typeof candidate.sourceName === 'string' ? candidate.sourceName : undefined,
				ordinal: typeof candidate.ordinal === 'number' && Number.isInteger(candidate.ordinal) && candidate.ordinal >= 0
					? candidate.ordinal
					: undefined,
				embed: typeof candidate.embed === 'boolean' ? candidate.embed : undefined,
			}]
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

export function readMissingAssets(frontMatter: Record<string, unknown>): MissingAsset[] {
	const raw = frontMatter['onenote-missing-assets'];
	if (!Array.isArray(raw)) return [];
	return raw.flatMap(value => {
		if (!value || typeof value !== 'object') return [];
		const candidate = value as Partial<MissingAsset>;
		return typeof candidate.name === 'string'
			&& typeof candidate.label === 'string'
			&& typeof candidate.embed === 'boolean'
			? [{
				name: candidate.name,
				label: candidate.label,
				embed: candidate.embed,
				sourceName: typeof candidate.sourceName === 'string' ? candidate.sourceName : undefined,
				ordinal: typeof candidate.ordinal === 'number' && Number.isInteger(candidate.ordinal) && candidate.ordinal >= 0
					? candidate.ordinal
					: undefined,
			}]
			: [];
	});
}

export function writeMissingAssets(frontMatter: Record<string, unknown>, missing: MissingAsset[]): void {
	if (missing.length > 0) frontMatter['onenote-missing-assets'] = missing;
	else delete frontMatter['onenote-missing-assets'];
}

function assetBaseName(asset: Asset): string {
	const name = path.posix.basename(asset.path);
	const extension = path.posix.extname(name);
	const base = extension ? name.slice(0, -extension.length) : name;
	return base.replace(/ (\d+)$/, '') + extension;
}

export function linkTargetForAsset(body: string, asset: Asset): { label: string, embed: boolean } | undefined {
	const encoded = encodeURI(asset.path);
	const at = body.indexOf(encoded);
	if (at < 0) return undefined;

	// Anchor on the destination's own `](`. OCR labels can contain unescaped
	// brackets, so searching backwards for the nearest `[` selects the wrong
	// boundary for labels such as "gene [source:HGNC ...".
	const labelEnd = body.lastIndexOf('](', at);
	if (labelEnd < 0) return undefined;
	const image = body.lastIndexOf('![', labelEnd);
	const link = body.lastIndexOf('[', labelEnd);
	const embed = image >= 0 && body.indexOf('](', image + 2) === labelEnd;
	const opening = embed ? image : link;
	if (opening < 0) return undefined;
	const labelStart = opening + (embed ? 2 : 1);
	return { label: body.slice(labelStart, labelEnd), embed };
}

/**
 * OCR labels can differ only because a prior conversion fenced short code
 * fragments inside an image label. Keep this deliberately strict: it is only
 * used after the same-page/name/embed checks below have already succeeded.
 */
export function sameAttachmentLabel(left: string, right: string): boolean {
	const normalize = (value: string) => value
		.replace(/^```[^\r\n]*\r?\n|^```\s*$/gmu, '')
		.replace(/\s+/gu, ' ')
		.trim()
		.toLocaleLowerCase();
	const first = normalize(left);
	const second = normalize(right);
	if (first === second) return true;
	if (first.length < 120 || second.length < 120 || first.length > 12_000 || second.length > 12_000) return false;
	const previous = Array.from({ length: second.length + 1 }, (_, index) => index);
	for (let row = 1; row <= first.length; row++) {
		let diagonal = previous[0];
		previous[0] = row;
		for (let column = 1; column <= second.length; column++) {
			const above = previous[column];
			previous[column] = Math.min(
				previous[column] + 1,
				previous[column - 1] + 1,
				diagonal + Number(first[row - 1] !== second[column - 1]),
			);
			diagonal = above;
		}
	}
	return previous[second.length] <= Math.max(3, Math.floor(Math.max(first.length, second.length) * 0.06));
}

/**
 * New conversions use source identity rather than rendered labels. Old staged
 * notes have neither field and retain the compatibility matcher below.
 */
export function attachmentIdentityMatches(asset: Asset, missing: MissingAsset): boolean {
	const identified = missing.sourceName !== undefined || missing.ordinal !== undefined || asset.sourceName !== undefined || asset.ordinal !== undefined;
	return identified
		? missing.sourceName === asset.sourceName
			&& missing.ordinal === asset.ordinal
			&& (asset.embed === undefined || missing.embed === asset.embed)
		: assetBaseName(asset) === missing.name;
}

function sourceAssetForMissing(note: PageNote, missing: MissingAsset, used: Set<string>): Asset | undefined {
	const candidates = readAssets(note.frontMatter).filter(asset => {
		const key = `${note.file}\u0000${asset.path}`;
		if (used.has(key)) return false;
		const link = linkTargetForAsset(note.body, asset);
		if (!link || link.embed !== missing.embed || !stageAsset(note, asset)) return false;

		// New conversions retain source identity. Never fall back to a visible
		// label when either side has that identity: OCR is intentionally absent
		// from current Markdown and is not evidence that two assets correspond.
		if (!attachmentIdentityMatches(asset, missing)) return false;
		const identified = missing.sourceName !== undefined || missing.ordinal !== undefined || asset.sourceName !== undefined || asset.ordinal !== undefined;
		return identified || sameAttachmentLabel(link.label, missing.label);
	});

	// Even exact source identity must not select an arbitrary duplicate from a
	// malformed or incomplete manifest.
	if (candidates.length !== 1) return undefined;
	const asset = candidates[0];
	used.add(`${note.file}\u0000${asset.path}`);
	return asset;
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
			sourceName: asset.sourceName,
			ordinal: asset.ordinal,
			embed: asset.embed,
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
		if (path.basename(file).toLowerCase() !== '_onenote archive.md' && pageId(note)) version.pages.push(note);
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

function linkKey(value: string): string | undefined {
	let decoded: string;
	try {
		decoded = decodeURI(value);
	}
	catch {
		return undefined;
	}
	decoded = posix(decoded).replace(/^\.\//, '').replace(/^\//, '').replace(/\.md$/iu, '');
	const parts = decoded.split('/');
	if (parts.some(part => part === '..')) return undefined;
	return normalized(parts.join('/'));
}

function splitLinkDestination(destination: string): { path: string, suffix: string } {
	const at = destination.search(/[?#]/);
	return at < 0 ? { path: destination, suffix: '' } : { path: destination.slice(0, at), suffix: destination.slice(at) };
}

function oneNoteTitle(destination: string): string | undefined {
	if (!destination.toLowerCase().startsWith('onenote:')) return undefined;
	const hash = destination.indexOf('#');
	if (hash < 0) return undefined;
	const tail = destination.slice(hash + 1);
	const encoded = tail.slice(0, tail.indexOf('&') < 0 ? tail.length : tail.indexOf('&'));
	try {
		let decoded = decodeURIComponent(encoded);
		if (/%[0-9a-f]{2}/i.test(decoded)) decoded = decodeURIComponent(decoded);
		return decoded;
	}
	catch {
		return undefined;
	}
}

function visibleLinkText(label: string, fallback: string): string {
	const visible = label.replace(/^!?\[/u, '').replace(/\]$/u, '').trim();
	return visible || fallback;
}

/** Rewrites only uniquely indexed local note links; external and unresolved links remain byte-for-byte unchanged. */
export function rewriteFinalLinks(
	body: string,
	sourceRelative: string,
	finalRelative: string,
	targets: Map<string, string[]>,
	crossNotebook: 'link' | 'title' = 'link',
): string {
	const fieldsRewritten = body.replace(/\uf7df?HYPERLINK "(onenote:[^"]+)"([^\r\n|]*)/giu, (whole, destination: string, visible: string) => {
		const linkedTitle = oneNoteTitle(destination);
		if (!linkedTitle) return whole;
		const matches = titleMatches(linkedTitle, targets);
		if (matches.size !== 1) return visible.trim() || linkedTitle;
		const [target] = matches;
		if (crossNotebook === 'title' && notebookComponent(finalRelative) !== notebookComponent(target)) return visible.trim() || linkedTitle;
		const rewritten = relativeFinalLink(finalRelative, target);
		const label = visible.trim() || linkedTitle;
		return `[${label}](${encodeURI(rewritten)})`;
	});

	return fieldsRewritten.replace(/(!?\[(?:\\.|[^\\\]\r\n])*\])\(((?:[^\s<>()]|\([^()\r\n]*\))+)(\s+(?:"[^"]*"|'[^']*'|\([^)]*\)))?\)(?: \*\(OneNote link target was not found in this import\)\*)?/g, (whole, label: string, destination: string, title = '') => {
		const linkedTitle = oneNoteTitle(destination);
		if (linkedTitle) {
			const matches = titleMatches(linkedTitle, targets);
			if (matches.size !== 1) return visibleLinkText(label, linkedTitle);
			const [target] = matches;
			if (crossNotebook === 'title' && notebookComponent(finalRelative) !== notebookComponent(target)) return visibleLinkText(label, linkedTitle);
			const rewritten = relativeFinalLink(finalRelative, target);
			return `${label}(${encodeURI(rewritten)}${title})`;
		}
		if (destination.startsWith('#') || /^[a-z][a-z\d+.-]*:/iu.test(destination)) return whole;
		const { path: destinationPath, suffix } = splitLinkDestination(destination);
		const direct = linkKey(destinationPath);
		const relative = linkKey(posix(path.posix.join(path.posix.dirname(sourceRelative), destinationPath)));
		const matches = new Set([...(direct ? targets.get(direct) ?? [] : []), ...(relative ? targets.get(relative) ?? [] : [])]);
		if (matches.size !== 1) return whole;
		const [target] = matches;
		const rewritten = relativeFinalLink(finalRelative, target);
		return `${label}(${encodeURI(rewritten)}${suffix}${title})`;
	});
}

function notebookComponent(relative: string): string {
	return posix(relative).split('/')[0] ?? '';
}

function relativeFinalLink(finalRelative: string, target: string): string {
	let rewritten = posix(path.posix.relative(path.posix.dirname(finalRelative), target)).replace(/\.md$/iu, '');
	if (!rewritten.startsWith('.')) rewritten = `./${rewritten}`;
	return rewritten;
}

function finalLinkIndex(plans: PlannedPage[], buildRoot: string): Map<string, string[]> {
	const index = new Map<string, string[]>();
	for (const plan of plans) {
		const finalRelative = posix(path.relative(buildRoot, plan.target));
		const titleCandidates = [plan.selected.frontMatter.title, path.basename(plan.target, path.extname(plan.target))];
		for (const title of titleCandidates) {
			if (typeof title !== 'string' || title === '') continue;
			for (const key of titleKeys(title)) {
				const matches = index.get(key) ?? [];
				if (!matches.includes(finalRelative)) matches.push(finalRelative);
				index.set(key, matches);
			}
		}
		for (const candidate of plan.candidates) {
			const key = linkKey(candidate.relative);
			if (!key) continue;
			const matches = index.get(key) ?? [];
			if (!matches.includes(finalRelative)) matches.push(finalRelative);
			index.set(key, matches);
		}
	}
	return index;
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
	const missing = readMissingAssets(frontMatter);
	const remainingMissing: MissingAsset[] = [];
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
	for (const missingAsset of missing) {
		const older = candidates
			.filter(candidate => candidate !== selected)
			.sort((left, right) => right.version.backupMtimeMs - left.version.backupMtimeMs);
		let result: CopiedAsset | undefined;
		for (const candidate of older) {
			const source = sourceAssetForMissing(candidate, missingAsset, used);
			if (!source) continue;
			result = copyAsset(candidate, source, target, buildRoot);
			if (result) break;
		}
		if (!result) {
			unrecovered++;
			remainingMissing.push(missingAsset);
			issues.push(`${selected.file}: could not recover missing ${missingAsset.name}`);
			continue;
		}
		copied.push(result.asset);
		recoveredLinks.push(markdownLink(missingAsset, result.markdownPath));
		recovered++;
	}

	frontMatter['onenote-assets'] = copied;
	writeMissingAssets(frontMatter, remainingMissing);
	frontMatter['onenote-completeness'] = unrecovered > 0 ? 'degraded' : recovered > 0 ? 'recovered' : frontMatter['onenote-completeness'];
	frontMatter['onenote-merged-from'] = selected.version.source;
	frontMatter['onenote-recovered-assets'] = recovered;
	body = appendRecoveredLinks(body, recoveredLinks);
	fs.mkdirSync(path.dirname(target), { recursive: true });
	fs.writeFileSync(target, serializeFrontMatter(frontMatter) + body, 'utf8');
	return { recovered, unrecovered };
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

function rewriteGlobalOneNoteLinks(outputRoot: string, notebookNames: string[]): void {
	const markdown: string[] = [];
	const walk = (directory: string): void => {
		for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
			const item = path.join(directory, entry.name);
			if (entry.isDirectory()) walk(item);
			else if (entry.isFile() && entry.name.toLowerCase().endsWith('.md')) markdown.push(item);
		}
	};
	for (const notebook of notebookNames) walk(path.join(outputRoot, notebook));

	const index = new Map<string, string[]>();
	for (const file of markdown) {
		const parsed = parseFrontMatterBlock(fs.readFileSync(file, 'utf8'));
		if (!parsed) continue;
		const relative = posix(path.relative(outputRoot, file));
		for (const title of [parsed.frontMatter.title, path.basename(file, path.extname(file))]) {
			if (typeof title !== 'string' || title === '') continue;
			for (const key of titleKeys(title)) {
				const matches = index.get(key) ?? [];
				if (!matches.includes(relative)) matches.push(relative);
				index.set(key, matches);
			}
		}
	}

	for (const file of markdown) {
		const parsed = parseFrontMatterBlock(fs.readFileSync(file, 'utf8'));
		if (!parsed) continue;
		const relative = posix(path.relative(outputRoot, file));
		const rewritten = rewriteFinalLinks(parsed.body, relative.replace(/\.md$/iu, ''), relative, index, 'title');
		if (rewritten !== parsed.body) fs.writeFileSync(file, serializeFrontMatter(parsed.frontMatter) + rewritten, 'utf8');
	}
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
		const plans: PlannedPage[] = [];

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
				plans.push({ selected, candidates: group.candidates, target });
				pages++;
			}
			report.sections.push({ name: sectionName, pages, recoveredAssets, unrecoveredAssets, skippedOlderOnlyPages });
		}

		const links = finalLinkIndex(plans, build);
		for (const plan of plans) {
			const section = report.sections.find(candidate => within(path.join(build, candidate.name), plan.target));
			const outcome = writePage(plan.selected, plan.candidates, plan.target, build, report.issues);
			if (section) {
				section.recoveredAssets += outcome.recovered;
				section.unrecoveredAssets += outcome.unrecovered;
			}
			const content = fs.readFileSync(plan.target, 'utf8');
			const parsed = parseFrontMatterBlock(content);
			if (!parsed) throw new Error(`Generated note has invalid frontmatter: ${plan.target}`);
			const sourceRelative = plan.selected.relative.replace(/\.md$/iu, '');
			const finalRelative = posix(path.relative(build, plan.target));
			const rewritten = rewriteFinalLinks(parsed.body, sourceRelative, finalRelative, links);
			fs.writeFileSync(plan.target, serializeFrontMatter(parsed.frontMatter) + rewritten, 'utf8');
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
	rewriteGlobalOneNoteLinks(outputRoot, notebookNames);
	for (const notebook of notebookNames) {
		const auditErrors = auditBuild(path.join(outputRoot, notebook));
		if (auditErrors.length > 0) throw new Error(`Refusing to publish global links for ${notebook}; ${auditErrors.join('\n')}`);
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
