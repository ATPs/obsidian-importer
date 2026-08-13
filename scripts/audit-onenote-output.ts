import '../tests/shims/runtime';

import * as crypto from 'node:crypto';
import * as fs from 'node:fs';
import * as path from 'node:path';

import { parseFrontMatterBlock } from '../src/util';

interface Asset {
	path: string;
	length: number;
	sha256: string;
}

interface Page {
	file: string;
	relativeFile: string;
	notebookRoot: string;
	body: string;
	frontMatter: Record<string, unknown>;
}

interface MergeNotebook {
	notebook?: unknown;
	sources?: unknown;
	sections?: unknown;
	issues?: unknown;
}

interface Finding {
	file: string;
	detail: string;
}

const [outputArgument, ...rawArguments] = process.argv.slice(2);
if (!outputArgument) throw new Error('Usage: audit-onenote-output.ts <output directory> [source directory ...]');
const previousArgument = rawArguments.find(argument => argument.startsWith('--previous='))?.slice('--previous='.length);
const sourceArguments = rawArguments.filter(argument => !argument.startsWith('--previous='));

const output = path.resolve(outputArgument);
if (!fs.statSync(output).isDirectory()) throw new Error(`Output is not a directory: ${output}`);

const normalizePath = (value: string): string => path.resolve(value).normalize().toLowerCase();
const sha256 = (bytes: Uint8Array): string => crypto.createHash('sha256').update(bytes).digest('hex');
const relativeOutputPath = (file: string): string => path.relative(output, file).replaceAll('\\', '/');

const sourceHashes = new Map<string, string>();
for (const argument of sourceArguments) {
	const source = path.resolve(argument);
	if (!fs.statSync(source).isDirectory()) throw new Error(`Source is not a directory: ${source}`);
	for (const entry of fs.readdirSync(source, { withFileTypes: true })) {
		if (!entry.isFile() || !['.one', '.onepkg', '.onex'].includes(path.extname(entry.name).toLowerCase())) continue;
		const file = path.join(source, entry.name);
		sourceHashes.set(normalizePath(file), sha256(fs.readFileSync(file)));
	}
}

const mergeReportFile = path.join(output, '_merge-report.json');
const mergeReport = fs.existsSync(mergeReportFile)
	? JSON.parse(fs.readFileSync(mergeReportFile, 'utf8')) as unknown
	: null;
const mergeNotebooks = Array.isArray(mergeReport) ? mergeReport as MergeNotebook[] : [];
const notebookNames = mergeNotebooks.flatMap(entry => typeof entry.notebook === 'string' ? [entry.notebook] : []);
const finalRoots = notebookNames.length > 0
	? notebookNames.map(name => path.join(output, name))
	: fs.readdirSync(output, { withFileTypes: true })
		.filter(entry => entry.isDirectory() && entry.name !== '_staging' && !entry.name.startsWith('.superseded-'))
		.map(entry => path.join(output, entry.name));

const rootEntries = fs.readdirSync(output, { withFileTypes: true });
const unexpectedRootEntries = rootEntries.flatMap(entry => {
	if (entry.name === '_merge-report.json' || notebookNames.includes(entry.name)) return [];
	return [{ name: entry.name, type: entry.isDirectory() ? 'directory' : 'file' }];
});

const files: string[] = [];
const mergeSourceHashes = new Map<string, string>();
for (const notebook of mergeNotebooks) {
	if (!Array.isArray(notebook.sources)) continue;
	for (const value of notebook.sources) {
		if (!value || typeof value !== 'object') continue;
		const source = value as { path?: unknown, sha256?: unknown };
		if (typeof source.path === 'string' && typeof source.sha256 === 'string') {
			mergeSourceHashes.set(normalizePath(source.path), source.sha256);
		}
	}
}
const walk = (directory: string): void => {
	if (!fs.existsSync(directory)) return;
	for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
		const file = path.join(directory, entry.name);
		if (entry.isDirectory()) walk(file);
		else if (entry.isFile()) files.push(file);
	}
};
for (const root of finalRoots) walk(root);

const markdown = files.filter(file => file.toLowerCase().endsWith('.md'));
const pageIds = new Map<string, string>();
const pages: Page[] = [];
const failures: Record<string, Finding[]> = {
	archiveFiles: [],
	codeFences: [],
	duplicatePageIds: [],
	frontmatter: [],
	flattenedLines: [],
	htmlFallbacks: [],
	localLinks: [],
	manifestAssets: [],
	attachmentConsistency: [],
	pageCounts: [],
	preservation: [],
	sourceHashes: [],
	unexpectedRootEntries: [],
};
const htmlFallbacks: { file: string, logicalFallbacks: number, tableTags: number, lines: number[], reason: string }[] = [];
const longestLines: { file: string, length: number, line: number }[] = [];
const flattenedParagraphCandidates: { file: string, length: number, line: number, kind: string }[] = [];
const shortPages: { file: string, characters: number }[] = [];
const missingAssetPages: { file: string, count: number, names: string[] }[] = [];
const recoveredAssetPages: { file: string, count: number }[] = [];
const manifestAssetDetails: { file: string, path: string, length: number, sha256: string }[] = [];
const manifestTargetsByPage = new Map<string, Set<string>>();
const allManifestTargets = new Set<string>();
const bodyAttachmentTargetsByPage = new Map<string, Set<string>>();
const allBodyAttachmentTargets = new Set<string>();
const retainedOneNoteLinkPages = new Map<string, number>();
let codeBlocks = 0;
let gfmTables = 0;
let htmlTables = 0;
let xmlFences = 0;
let preservationHeadings = 0;
let manifestAssets = 0;
let recoveredAssets = 0;
let missingAssets = 0;
const changedSourcesAfterConversion: { path: string, recordedSha256: string, currentSha256: string }[] = [];
for (const [source, currentSha256] of sourceHashes) {
	const recordedSha256 = mergeSourceHashes.get(source);
	if (recordedSha256 && recordedSha256 !== currentSha256) {
		changedSourcesAfterConversion.push({ path: source, recordedSha256, currentSha256 });
	}
}

const addFailure = (category: keyof typeof failures, file: string, detail: string): void => {
	failures[category].push({ file: relativeOutputPath(file), detail });
};
for (const entry of unexpectedRootEntries) {
	addFailure('unexpectedRootEntries', path.join(output, entry.name), `Unexpected root ${entry.type}; final output must contain only notebooks and _merge-report.json`);
}

function assets(frontMatter: Record<string, unknown>, file: string): Asset[] {
	const raw = frontMatter['onenote-assets'];
	if (!Array.isArray(raw)) {
		addFailure('manifestAssets', file, 'onenote-assets must be an array');
		return [];
	}
	return raw.flatMap(value => {
		if (!value || typeof value !== 'object') {
			addFailure('manifestAssets', file, 'Malformed non-object asset manifest entry');
			return [];
		}
		const asset = value as Partial<Asset>;
		return typeof asset.path === 'string' && typeof asset.length === 'number' && typeof asset.sha256 === 'string'
			? [asset as Asset]
			: (addFailure('manifestAssets', file, `Malformed asset manifest entry: ${JSON.stringify(value)}`), []);
	});
}

function resolveManifestAsset(page: Page, assetPath: string): string | null {
	const target = path.resolve(page.notebookRoot, ...assetPath.replaceAll('\\', '/').split('/'));
	const relative = path.relative(page.notebookRoot, target);
	return relative.startsWith('..') || path.isAbsolute(relative) ? null : target;
}

for (const file of markdown) {
	if (path.basename(file).toLowerCase() === '_onenote archive.md') {
		addFailure('archiveFiles', file, 'archive note must not be generated in the merged output');
	}
	const content = fs.readFileSync(file, 'utf8');
	const parsed = parseFrontMatterBlock(content);
	if (!parsed || typeof parsed.frontMatter['onenote-id'] !== 'string') {
		addFailure('frontmatter', file, 'Markdown file has no valid onenote-id');
		continue;
	}
	const relativeFile = relativeOutputPath(file);
	const notebookComponent = relativeFile.split('/')[0];
	const page: Page = {
		file,
		relativeFile,
		notebookRoot: path.join(output, notebookComponent),
		body: parsed.body,
		frontMatter: parsed.frontMatter,
	};
	pages.push(page);

	const id = parsed.frontMatter['onenote-id'];
	const duplicate = pageIds.get(id);
	if (duplicate) addFailure('duplicatePageIds', file, `Page ID ${id} is also used by ${relativeOutputPath(duplicate)}`);
	else pageIds.set(id, file);

	if (sourceHashes.size > 0) {
		const source = parsed.frontMatter['onenote-source'];
		const hash = parsed.frontMatter['onenote-source-sha256'];
		if (typeof source !== 'string') addFailure('sourceHashes', file, 'onenote-source is missing');
		else if (!sourceHashes.has(normalizePath(source))) addFailure('sourceHashes', file, `Source is not one of the selected inputs: ${source}`);
		else if (typeof hash !== 'string' || sourceHashes.get(normalizePath(source)) !== hash) {
			addFailure('sourceHashes', file, `Source SHA-256 mismatch for ${source}: recorded ${String(hash)}, current ${String(sourceHashes.get(normalizePath(source)))}`);
		}
	}

	const pageManifestTargets = new Set<string>();
	manifestTargetsByPage.set(normalizePath(file), pageManifestTargets);
	for (const asset of assets(parsed.frontMatter, file)) {
		manifestAssets++;
		const target = resolveManifestAsset(page, asset.path);
		if (!target) {
			addFailure('manifestAssets', file, `Asset escapes notebook root: ${asset.path}`);
			continue;
		}
		const normalizedTarget = normalizePath(target);
		pageManifestTargets.add(normalizedTarget);
		allManifestTargets.add(normalizedTarget);
		if (!fs.existsSync(target) || !fs.statSync(target).isFile()) {
			addFailure('manifestAssets', file, `Missing asset: ${asset.path}`);
			continue;
		}
		const bytes = fs.readFileSync(target);
		if (bytes.length !== asset.length) addFailure('manifestAssets', file, `Wrong asset length: ${asset.path}`);
		if (sha256(bytes) !== asset.sha256) addFailure('manifestAssets', file, `Wrong asset SHA-256: ${asset.path}`);
		manifestAssetDetails.push({ file: relativeFile, path: asset.path, length: bytes.length, sha256: sha256(bytes) });
	}

	const missing = parsed.frontMatter['onenote-missing-assets'];
	if (Array.isArray(missing) && missing.length > 0) {
		missingAssets += missing.length;
		const names = missing.flatMap(value => {
			if (!value || typeof value !== 'object') return [];
			const name = (value as { name?: unknown }).name;
			return typeof name === 'string' ? [name] : [];
		});
		missingAssetPages.push({ file: relativeFile, count: missing.length, names });
	}
	const recovered = parsed.frontMatter['onenote-recovered-assets'];
	if (typeof recovered === 'number' && recovered > 0) {
		recoveredAssets += recovered;
		recoveredAssetPages.push({ file: relativeFile, count: recovered });
	}

	const lines = parsed.body.split(/\r?\n/);
	let insideFence = false;
	let longest = 0;
	let longestAt = 0;
	const htmlLines: number[] = [];
	for (const [index, line] of lines.entries()) {
		if (/^```/.test(line)) {
			if (!insideFence) codeBlocks++;
			insideFence = !insideFence;
			continue;
		}
		if (insideFence) continue;
		if (line.length > longest) {
			longest = line.length;
			longestAt = index + 1;
		}
		if (line.length >= 1000) {
			const kind = /^\|.*\|$/u.test(line) ? 'gfm-table-row'
				: /^!?\[[^\]]*\]\(/u.test(line) ? 'link-or-image'
					: /^\s*(?:[-*+] |\d+[.)] )/u.test(line) ? 'list-item'
						: 'paragraph';
			flattenedParagraphCandidates.push({ file: relativeFile, length: line.length, line: index + 1, kind });
			const formattingNoise = (line.match(/==[🔴🟠🟡🟢🔵🟣]/gu) ?? []).length;
			if (formattingNoise >= 5 || (kind === 'gfm-table-row' && line.length >= 10_000)) {
				addFailure('flattenedLines', file, `${kind} has ${line.length} characters and conversion noise at body line ${index + 1}`);
			}
		}
		if (/<table(?:\s|>)/iu.test(line)) htmlLines.push(index + 1);
	}
	if (insideFence) addFailure('codeFences', file, 'Unclosed fenced code block');
	longestLines.push({ file: relativeFile, length: longest, line: longestAt });
	gfmTables += (parsed.body.match(/^\|.*\|\r?\n\|(?:\s*:?-+:?\s*\|)+/gmu) ?? []).length;
	const htmlCount = (parsed.body.match(/<table(?:\s|>)/giu) ?? []).length;
	htmlTables += htmlCount;
	if (htmlCount > 0) {
		const declared = parsed.frontMatter['onenote-html-fallbacks'];
		htmlFallbacks.push({
			file: relativeFile,
			logicalFallbacks: typeof declared === 'number' ? declared : 0,
			tableTags: htmlCount,
			lines: htmlLines,
			reason: 'Nested two-dimensional table cannot be represented safely as GFM or a parent-child list',
		});
		if (declared !== 1) addFailure('htmlFallbacks', file, `${htmlCount} HTML table(s) but frontmatter declares ${String(declared)}`);
	}
	const pageXmlFences = (parsed.body.match(/^```xml\s*$/gmu) ?? []).length;
	const pagePreservationHeadings = (parsed.body.match(/^## OneNote preservation data\s*$/gmu) ?? []).length;
	xmlFences += pageXmlFences;
	preservationHeadings += pagePreservationHeadings;
	if (pageXmlFences > 0) addFailure('preservation', file, `${pageXmlFences} XML fence(s)`);
	if (pagePreservationHeadings > 0) addFailure('preservation', file, `${pagePreservationHeadings} preservation heading(s)`);

	const plainCharacters = parsed.body.replace(/\s/g, '').length;
	if (plainCharacters < 40) shortPages.push({ file: relativeFile, characters: plainCharacters });
}

const markdownByAbsolutePath = new Set(markdown.map(normalizePath));
const markdownByStem = new Map<string, string[]>();
for (const file of markdown) {
	const stem = path.basename(file, path.extname(file)).normalize('NFC').toLowerCase();
	markdownByStem.set(stem, [...(markdownByStem.get(stem) ?? []), file]);
}

function oneNotePageTitle(url: string): string | null {
	const hash = url.indexOf('#');
	if (hash < 0) return null;
	const tail = url.slice(hash + 1);
	const encoded = tail.slice(0, Math.max(0, tail.indexOf('&') < 0 ? tail.length : tail.indexOf('&')));
	if (!encoded) return null;
	try {
		let decoded = decodeURIComponent(encoded);
		if (/%[0-9a-f]{2}/i.test(decoded)) decoded = decodeURIComponent(decoded);
		return decoded;
	}
	catch {
		return null;
	}
}

function stripMarkdownDestination(raw: string): string {
	const trimmed = raw.trim();
	if (trimmed.startsWith('<')) {
		const close = trimmed.indexOf('>');
		if (close >= 0) return trimmed.slice(1, close);
	}
	const title = trimmed.match(/^(.*?)(?:\s+["'].*["']|\s+\(.*\))$/u);
	return (title?.[1] ?? trimmed).trim();
}

function inlineDestinations(markdownBody: string): string[] {
	const destinations: string[] = [];
	for (let index = 0; index < markdownBody.length; index++) {
		if (markdownBody[index] !== ']' || markdownBody[index + 1] !== '(') continue;
		let depth = 1;
		let escaped = false;
		for (let cursor = index + 2; cursor < markdownBody.length; cursor++) {
			const character = markdownBody[cursor];
			if (escaped) {
				escaped = false;
				continue;
			}
			if (character === '\\') {
				escaped = true;
				continue;
			}
			if (character === '(') depth++;
			else if (character === ')' && --depth === 0) {
				destinations.push(markdownBody.slice(index + 2, cursor));
				index = cursor;
				break;
			}
		}
	}
	return destinations;
}

function withoutFencedCode(markdownBody: string): string {
	let insideFence = false;
	return markdownBody.split(/(\r?\n)/).map(part => {
		if (/^```/.test(part)) {
			insideFence = !insideFence;
			return '';
		}
		return insideFence && !/^\r?\n$/.test(part) ? ' '.repeat(part.length) : part;
	}).join('');
}

function localTargetExists(page: Page, rawDestination: string): boolean {
	let destination = stripMarkdownDestination(rawDestination).split('#')[0].split('?')[0];
	if (!destination) return true;
	try {
		destination = decodeURI(destination);
	}
	catch {
		return false;
	}
	const target = path.resolve(path.dirname(page.file), ...destination.replaceAll('\\', '/').split('/'));
	const relative = path.relative(output, target);
	if (relative.startsWith('..') || path.isAbsolute(relative)) return false;
	return fs.existsSync(target) || markdownByAbsolutePath.has(normalizePath(`${target}.md`));
}

function localExistingTarget(page: Page, rawDestination: string): string | null {
	let destination = stripMarkdownDestination(rawDestination).split('#')[0].split('?')[0];
	if (!destination) return null;
	try {
		destination = decodeURI(destination);
	}
	catch {
		return null;
	}
	const target = path.resolve(path.dirname(page.file), ...destination.replaceAll('\\', '/').split('/'));
	const relative = path.relative(output, target);
	return relative.startsWith('..') || path.isAbsolute(relative) || !fs.existsSync(target) || !fs.statSync(target).isFile()
		? null
		: target;
}

for (const page of pages) {
	const linkBody = withoutFencedCode(page.body);
	const bodyAttachmentTargets = new Set<string>();
	bodyAttachmentTargetsByPage.set(normalizePath(page.file), bodyAttachmentTargets);
	const retainedOneNoteLinks = linkBody.match(/onenote:[^\s)>]+/giu) ?? [];
	if (retainedOneNoteLinks.length > 0) retainedOneNoteLinkPages.set(page.relativeFile, retainedOneNoteLinks.length);
	for (const url of retainedOneNoteLinks) {
		const title = oneNotePageTitle(url);
		if (!title) continue;
		const matches = markdownByStem.get(title.normalize('NFC').toLowerCase()) ?? [];
		if (matches.length > 0) addFailure('localLinks', page.file, `Retained OneNote link has ${matches.length} imported title match(es): ${title}`);
	}
	const seenFindings = new Set<string>();
	for (const rawDestination of inlineDestinations(linkBody)) {
		const destination = stripMarkdownDestination(rawDestination);
		if (/^(?:https?|ftp|file|onenote|onemore|mailto|javascript|data):/iu.test(destination) || destination.startsWith('#')) continue;
		const existing = localExistingTarget(page, rawDestination);
		if (existing && path.extname(existing).toLowerCase() !== '.md') {
			const normalizedTarget = normalizePath(existing);
			bodyAttachmentTargets.add(normalizedTarget);
			allBodyAttachmentTargets.add(normalizedTarget);
		}
		const detail = `Unresolved Markdown link: ${destination}`;
		if (!localTargetExists(page, rawDestination) && !seenFindings.has(detail)) {
			seenFindings.add(detail);
			addFailure('localLinks', page.file, detail);
		}
	}
	for (const match of linkBody.matchAll(/!?\[\[([^\]\r\n]+)\]\]/gu)) {
		const targetText = match[1].split('|')[0].split('#')[0].trim();
		if (!targetText) continue;
		let decoded = targetText;
		try {
			decoded = decodeURI(targetText);
		}
		catch {
			const detail = `Malformed wiki link: ${targetText}`;
			if (!seenFindings.has(detail)) {
				seenFindings.add(detail);
				addFailure('localLinks', page.file, detail);
			}
			continue;
		}
		const relativeTarget = path.resolve(path.dirname(page.file), ...decoded.replaceAll('\\', '/').split('/'));
		const stemMatches = markdownByStem.get(path.basename(decoded, path.extname(decoded)).normalize('NFC').toLowerCase()) ?? [];
		if (!fs.existsSync(relativeTarget) && !markdownByAbsolutePath.has(normalizePath(`${relativeTarget}.md`)) && stemMatches.length === 0) {
			const detail = `Unresolved wiki link: ${targetText}`;
			if (!seenFindings.has(detail)) {
				seenFindings.add(detail);
				addFailure('localLinks', page.file, detail);
			}
		}
	}
}

for (const page of pages) {
	const manifest = manifestTargetsByPage.get(normalizePath(page.file)) ?? new Set<string>();
	const body = bodyAttachmentTargetsByPage.get(normalizePath(page.file)) ?? new Set<string>();
	for (const target of body) {
		if (!manifest.has(target)) addFailure('attachmentConsistency', page.file, `Body attachment is absent from manifest: ${relativeOutputPath(target)}`);
	}
	for (const target of manifest) {
		if (!body.has(target)) addFailure('attachmentConsistency', page.file, `Manifest attachment is not referenced by the page body: ${relativeOutputPath(target)}`);
	}
}
const diskAttachmentTargets = new Set(files
	.filter(file => !file.toLowerCase().endsWith('.md') && path.basename(file).toLowerCase() !== '_merge-report.json')
	.map(normalizePath));
for (const target of diskAttachmentTargets) {
	if (!allManifestTargets.has(target)) addFailure('attachmentConsistency', target, 'Attachment file exists on disk but is absent from every manifest');
}
for (const target of allManifestTargets) {
	if (!diskAttachmentTargets.has(target)) addFailure('attachmentConsistency', target, 'Manifest attachment is absent from the notebook file tree');
}

let expectedPages: number | null = null;
let mergeSourceCount: number | null = null;
const mergeIssues: string[] = [];
if (mergeNotebooks.length > 0) {
	expectedPages = 0;
	mergeSourceCount = 0;
	for (const notebook of mergeNotebooks) {
		if (Array.isArray(notebook.sources)) mergeSourceCount += notebook.sources.length;
		if (Array.isArray(notebook.issues)) mergeIssues.push(...notebook.issues.filter((issue): issue is string => typeof issue === 'string'));
		if (!Array.isArray(notebook.sections)) continue;
		for (const section of notebook.sections) {
			if (!section || typeof section !== 'object') continue;
			const count = (section as { pages?: unknown }).pages;
			if (typeof count === 'number') expectedPages += count;
		}
	}
	if (pages.length !== expectedPages) addFailure('pageCounts', mergeReportFile, `Found ${pages.length} pages; merge report declares ${expectedPages}`);
	if (sourceHashes.size > 0 && mergeSourceCount !== sourceHashes.size) {
		addFailure('pageCounts', mergeReportFile, `Merge report has ${mergeSourceCount} sources; selected inputs contain ${sourceHashes.size}`);
	}
}

const failureCount = Object.values(failures).reduce((sum, findings) => sum + findings.length, 0);
const report = {
	output,
	notebooks: finalRoots.map(root => path.basename(root)),
	unexpectedRootEntries,
	selectedSourceFiles: sourceHashes.size,
	mergeReportSourceFiles: mergeSourceCount,
	markdownFiles: markdown.length,
	pages: pages.length,
	expectedPages,
	uniquePageIds: pageIds.size,
	codeBlocks,
	gfmTables,
	htmlTables,
	htmlFallbacks,
	xmlFences,
	preservationHeadings,
	archiveFiles: failures.archiveFiles.length,
	manifestAssets,
	bodyAttachmentReferences: allBodyAttachmentTargets.size,
	diskAttachmentFiles: diskAttachmentTargets.size,
	manifestAssetDetails: manifestAssetDetails.sort((left, right) => left.file.localeCompare(right.file) || left.path.localeCompare(right.path)),
	recoveredAssets,
	recoveredAssetPages: recoveredAssetPages.sort((left, right) => right.count - left.count),
	missingAssets,
	missingAssetPages: missingAssetPages.sort((left, right) => right.count - left.count),
	mergeIssues,
	retainedOneNoteLinks: [...retainedOneNoteLinkPages.values()].reduce((sum, count) => sum + count, 0),
	retainedOneNoteLinkPages: [...retainedOneNoteLinkPages].map(([file, count]) => ({ file, count })),
	changedSourcesAfterConversion,
	longestNonCodeLines: longestLines.sort((left, right) => right.length - left.length).slice(0, 50),
	flattenedParagraphCandidates: flattenedParagraphCandidates.sort((left, right) => right.length - left.length),
	shortPages: shortPages.sort((left, right) => left.characters - right.characters),
	shortPageCount: shortPages.length,
	failureCount,
	failures,
};

const previous = previousArgument && fs.existsSync(previousArgument)
	? JSON.parse(fs.readFileSync(previousArgument, 'utf8')) as Record<string, unknown>
	: null;
const comparisonKeys = ['markdownFiles', 'pages', 'uniquePageIds', 'codeBlocks', 'gfmTables', 'htmlTables', 'xmlFences', 'preservationHeadings', 'archiveFiles', 'manifestAssets', 'recoveredAssets', 'missingAssets', 'failureCount'] as const;
const previousComparison = previous ? Object.fromEntries(comparisonKeys.map(key => {
	const current = report[key];
	const prior = previous[key];
	return [key, { previous: prior, current, delta: typeof current === 'number' && typeof prior === 'number' ? current - prior : null }];
})) : null;

process.stdout.write(JSON.stringify({ ...report, previousComparison }, null, 2) + '\n');
if (failureCount > 0) process.exitCode = 1;
