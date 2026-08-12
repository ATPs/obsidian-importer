import '../tests/shims/runtime';

import * as crypto from 'node:crypto';
import * as fs from 'node:fs';
import * as path from 'node:path';
import sax from 'sax';

import { parseFrontMatterBlock } from '../src/util';

const [sourceArgument, outputArgument] = process.argv.slice(2);
if (!sourceArgument || !outputArgument) throw new Error('Give source and output directories.');

const source = path.resolve(sourceArgument);
const output = path.resolve(outputArgument);
const sourceHashes = new Map<string, string>();
for (const entry of fs.readdirSync(source, { withFileTypes: true })) {
	if (!entry.isFile() || path.extname(entry.name).toLowerCase() !== '.one') continue;
	const fullPath = path.join(source, entry.name);
	sourceHashes.set(fullPath.toLowerCase(), crypto.createHash('sha256').update(fs.readFileSync(fullPath)).digest('hex'));
}

const markdown: string[] = [];
const walk = (directory: string): void => {
	for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
		const fullPath = path.join(directory, entry.name);
		if (entry.isDirectory()) walk(fullPath);
		else if (entry.isFile() && entry.name.toLowerCase().endsWith('.md')) markdown.push(fullPath);
	}
};
walk(output);

let xmlBlocks = 0;
let manifests = 0;
let assets = 0;
let missingPayloadRecords = 0;
const errors: string[] = [];

for (const file of markdown) {
	const content = fs.readFileSync(file, 'utf8');
	const parsed = parseFrontMatterBlock(content);
	if (!parsed) {
		// A OneNote attachment may itself be Markdown. Imported notes always have
		// OneNote frontmatter; attachment bytes are audited through their owner.
		continue;
	}

	const recordedSource = parsed.frontMatter['onenote-source'];
	const recordedHash = parsed.frontMatter['onenote-source-sha256'];
	if (typeof recordedSource !== 'string' || sourceHashes.get(path.resolve(recordedSource).toLowerCase()) !== recordedHash) {
		errors.push(`${file}: source path/hash does not match an input file`);
	}

	const manifest = parsed.frontMatter['onenote-assets'];
	if (Array.isArray(manifest)) {
		manifests++;
		for (const value of manifest) {
			assets++;
			if (!value || typeof value !== 'object') {
				errors.push(`${file}: malformed asset manifest entry`);
				continue;
			}
			const asset = value as { path?: unknown, length?: unknown, sha256?: unknown };
			if (typeof asset.path !== 'string' || typeof asset.length !== 'number' || typeof asset.sha256 !== 'string') {
				errors.push(`${file}: incomplete asset manifest entry`);
				continue;
			}
			const target = path.resolve(output, ...asset.path.replaceAll('\\', '/').split('/'));
			if (!fs.existsSync(target)) {
				errors.push(`${file}: missing asset ${asset.path}`);
				continue;
			}
			const bytes = fs.readFileSync(target);
			if (bytes.length !== asset.length) errors.push(`${file}: wrong asset length ${asset.path}`);
			if (crypto.createHash('sha256').update(bytes).digest('hex') !== asset.sha256) errors.push(`${file}: wrong asset hash ${asset.path}`);
		}
	}

	for (const match of content.matchAll(/```xml\r?\n([\s\S]*?)\r?\n```/g)) {
		xmlBlocks++;
		missingPayloadRecords += (match[1].match(/code="ONENOTE_ASSET_DATA_MISSING"/g) ?? []).length;
		const parser = sax.parser(true);
		let parseError: Error | undefined;
		parser.onerror = error => { parseError = error; };
		parser.write(match[1]).close();
		if (parseError) errors.push(`${file}: invalid preservation XML: ${parseError.message}`);
	}
}

process.stdout.write(JSON.stringify({
	markdown: markdown.length,
	xmlBlocks,
	manifests,
	manifestAssets: assets,
	missingPayloadRecords,
	sourceFiles: sourceHashes.size,
	errors,
}, null, 2) + '\n');
if (errors.length > 0) process.exitCode = 1;
