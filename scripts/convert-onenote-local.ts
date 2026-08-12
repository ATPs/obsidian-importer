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
import { MemoryVault, memoryApp } from '../tests/shims/vault';

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
