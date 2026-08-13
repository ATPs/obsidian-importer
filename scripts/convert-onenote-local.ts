import '../tests/shims/runtime';

import * as nodeCrypto from 'node:crypto';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import * as url from 'node:url';
import * as zlib from 'node:zlib';

import { NodePickedFile, provideNodeModules } from '../src/filesystem';
import { OneNoteFileImporter } from '../src/formats/onenote-file';
import { ImportContext } from '../src/import-context';
import { extensionFromBytes, parseFrontMatterBlock, serializeFrontMatter } from '../src/util';
import { MemoryVault, memoryApp } from '../tests/shims/vault';
import { convertEmfToPng, emfPngSettings } from './emf-to-png';

class ConsoleContext extends ImportContext {
	protected onStatus(message: string): void {
		process.stderr.write(`[status] ${message}\n`);
	}

	protected onNoteSuccess(name: string): void {
		process.stderr.write(`[note ${this.notes}] ${name}\n`);
	}

	protected onAttachmentSuccess(name: string): void {
		if (this.attachments % 100 === 0) process.stderr.write(`[attachments] ${this.attachments}: ${name}\n`);
	}

	protected onProgress(current: number, total: number): void {
		if (current === 1 || current === total || current % 25 === 0) process.stderr.write(`[pages] ${current}/${total}\n`);
	}
}

export interface LocalConversionReport {
	source: string;
	destination: string;
	inputFiles: string[];
	notes: number;
	attachments: number;
	failed: unknown[];
	skipped: unknown[];
	log: unknown[];
	outputFiles: string[];
	emfConversions: EmfConversion[];
}

export interface EmfConversion {
	input: string;
	output: string;
	width: number;
	height: number;
}

interface AssetManifestEntry {
	path: string;
	length: number;
	sha256: string;
	sourceName?: string;
	ordinal?: number;
	embed?: boolean;
}

function utf8(content: string | ArrayBuffer): string {
	return typeof content === 'string' ? content : new TextDecoder().decode(content);
}

function replaceAssetReference(content: string, before: string, after: string): string {
	return content.replaceAll(encodeURI(before), encodeURI(after)).replaceAll(before, after);
}

/**
 * The portable backup export intentionally replaces EMF with PNG for Obsidian.
 * The interactive importer does not use this function and continues preserving
 * original EMF bytes.
 */
async function replaceEmfWithPng(vault: MemoryVault): Promise<EmfConversion[]> {
	const conversions: EmfConversion[] = [];
	const emfPaths = vault.paths().filter(file => file.toLocaleLowerCase().endsWith('.emf'));

	for (const emfPath of emfPaths) {
		const source = vault.contents.get(emfPath);
		if (!(source instanceof ArrayBuffer)) throw new Error(`EMF attachment is not binary: ${emfPath}`);
		const png = await convertEmfToPng(new Uint8Array(source));
		const base = emfPath.slice(0, -4);
		let pngPath = `${base}.png`;
		// A OneNote page can contain a PNG and EMF that share the generated
		// attachment stem. Keep both images instead of overwriting the PNG.
		if (vault.contents.has(pngPath)) {
			for (let index = 0; ; index++) {
				const suffix = index === 0 ? ' EMF' : ` EMF ${index}`;
				pngPath = `${base}${suffix}.png`;
				if (!vault.contents.has(pngPath)) break;
			}
		}
		vault.contents.set(pngPath, png.bytes.buffer.slice(png.bytes.byteOffset, png.bytes.byteOffset + png.bytes.byteLength));
		vault.remove(emfPath);
		conversions.push({ input: emfPath, output: pngPath, width: png.width, height: png.height });
	}

	const renamed = new Map<string, string>();
	for (const assetPath of vault.paths()) {
		const content = vault.contents.get(assetPath);
		if (!(content instanceof ArrayBuffer) || assetPath.toLocaleLowerCase().endsWith('.emf')) continue;
		const extension = extensionFromBytes(new Uint8Array(content));
		const current = assetPath.match(/\.[^.\\/]+$/u)?.[0].slice(1).toLocaleLowerCase();
		const compatible = extension === 'jpg' ? current === 'jpg' || current === 'jpeg' : current === extension;
		if (!extension || compatible) continue;
		const next = assetPath.replace(/\.[^.\\/]+$/u, `.${extension}`);
		if (vault.contents.has(next)) throw new Error(`Attachment type correction would overwrite an asset: ${next}`);
		vault.contents.set(next, content);
		vault.remove(assetPath);
		renamed.set(assetPath, next);
	}

	if (conversions.length === 0 && renamed.size === 0) return conversions;
	const converted = new Map(conversions.map(item => [item.input, item]));
	for (const markdownPath of vault.paths().filter(file => file.toLocaleLowerCase().endsWith('.md'))) {
		const content = vault.contents.get(markdownPath);
		if (!content) continue;
		const parsed = parseFrontMatterBlock(utf8(content));
		if (!parsed) throw new Error(`Cannot update EMF manifest without frontmatter: ${markdownPath}`);
		const rawAssets = parsed.frontMatter['onenote-assets'];
		if (!Array.isArray(rawAssets)) continue;
		let changed = false;
		const labelReplacements: { before: string, after: string }[] = [];
		const assets: unknown[] = [];
		for (const value of rawAssets) {
			if (!value || typeof value !== 'object') {
				assets.push(value);
				continue;
			}
			const asset = value as AssetManifestEntry;
			const vaultPath = asset.path.replaceAll('\\', '/');
			const replacement = converted.get(vaultPath);
			const renamedPath = renamed.get(vaultPath);
			if (!replacement && !renamedPath) {
				assets.push(value);
				continue;
			}
			changed = true;
			const output = replacement?.output ?? renamedPath!;
			const bytes = vault.contents.get(output);
			if (!(bytes instanceof ArrayBuffer)) throw new Error(`Converted PNG was not written: ${output}`);
			const sourceName = replacement
				? asset.sourceName?.replace(/\.emf$/iu, ' (EMF converted).png')
				: asset.sourceName?.replace(/\.[^.\\/]+$/u, `.${path.extname(output).slice(1)}`);
			if (asset.sourceName && sourceName && asset.sourceName !== sourceName) {
				labelReplacements.push({ before: asset.sourceName, after: sourceName });
			}
			assets.push({
				...asset,
				path: output,
				length: bytes.byteLength,
				sha256: await sha256(bytes),
				sourceName,
			});
		}
		if (!changed) continue;
		parsed.frontMatter['onenote-assets'] = assets;
		let body = parsed.body;
		for (const [before, item] of converted) body = replaceAssetReference(body, before, item.output);
		for (const [before, after] of renamed) body = replaceAssetReference(body, before, after);
		for (const label of labelReplacements) body = body.replaceAll(`![${label.before}]`, `![${label.after}]`);
		vault.contents.set(markdownPath, serializeFrontMatter(parsed.frontMatter) + body);
	}

	return conversions;
}

async function sha256(bytes: ArrayBuffer): Promise<string> {
	return nodeCrypto.createHash('sha256').update(new Uint8Array(bytes)).digest('hex');
}

/**
 * Run the existing OneNote file importer outside Obsidian. The output vault is
 * configured like a portable vault: every note owns an attachments subfolder.
 */
export async function convertOneNoteLocal(sourceArgument: string, destinationArgument: string): Promise<LocalConversionReport> {
	provideNodeModules({ nodeCrypto, fs, os, path, url, zlib });

	const source = path.resolve(sourceArgument);
	const destination = path.resolve(destinationArgument);
	const sourceStat = fs.statSync(source);
	const files = sourceStat.isFile()
		? [source]
		: fs.readdirSync(source, { withFileTypes: true })
			.filter(entry => entry.isFile() && ['.one', '.onepkg', '.onex'].includes(path.extname(entry.name).toLowerCase()))
			.map(entry => path.join(source, entry.name));

	if (files.length === 0) throw new Error(`No OneNote files found in ${source}`);
	if (fs.existsSync(destination)) throw new Error(`Destination already exists: ${destination}`);

	const vault = new MemoryVault();
	vault.config.set('attachmentFolderPath', './attachments');
	const subject = new OneNoteFileImporter(memoryApp(vault), {
		sourceEl: null,
		outputEl: null,
		optionsEl: null,
	} as never);
	await subject.ready;
	subject.files = files.map(file => new NodePickedFile(file));
	subject.outputLocation = '/';
	subject.indexImportedNotes();

	const ctx = new ConsoleContext();
	await subject.import(ctx);
	await subject.finalizeMarkdownOutput(ctx);
	const emfConversions = await replaceEmfWithPng(vault);

	fs.mkdirSync(destination, { recursive: false });
	for (const [vaultPath, content] of vault.contents) {
		const relative = vaultPath.replaceAll('\\', '/').replace(/^\/+/, '');
		const target = path.resolve(destination, ...relative.split('/'));
		const withinDestination = path.relative(destination, target);
		if (withinDestination.startsWith('..') || path.isAbsolute(withinDestination)) {
			throw new Error(`Refusing to write outside destination: ${vaultPath}`);
		}

		fs.mkdirSync(path.dirname(target), { recursive: true });
		if (typeof content === 'string') fs.writeFileSync(target, content, 'utf8');
		else fs.writeFileSync(target, new Uint8Array(content));
	}

	const report: LocalConversionReport = {
		source,
		destination,
		inputFiles: files,
		notes: ctx.notes,
		attachments: ctx.attachments,
		failed: ctx.failed,
		skipped: ctx.skipped,
		log: ctx.log.map(entry => ({
			...entry,
			reason: entry.reason instanceof Error ? entry.reason.message : entry.reason,
		})),
		outputFiles: vault.paths(),
		emfConversions,
	};
	fs.writeFileSync(path.join(destination, '_conversion-report.json'), JSON.stringify(report, null, 2), 'utf8');
	return report;
}

async function main(): Promise<void> {
	const [sourceArgument, destinationArgument] = process.argv.slice(2);
	if (!sourceArgument || !destinationArgument) {
		throw new Error('Usage: convert-onenote-local.ts <source .one file or directory> <destination directory>');
	}

	const report = await convertOneNoteLocal(sourceArgument, destinationArgument);

	process.stdout.write(JSON.stringify({
		inputFiles: report.inputFiles.length,
		outputFiles: report.outputFiles.length,
		notes: report.notes,
		attachments: report.attachments,
		failed: report.failed.length,
		skipped: report.skipped.length,
	}, null, 2) + '\n');
}

if (process.argv[1] && path.resolve(process.argv[1]) === path.resolve(url.fileURLToPath(import.meta.url))) {
	void main().catch(error => {
		console.error(error);
		process.exitCode = 1;
	});
}
