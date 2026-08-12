import '../shims/runtime';

import { test } from 'node:test';
import assert from 'node:assert/strict';
import type { TFile, TFolder } from 'obsidian';

import { OneNoteFileImporter } from '../../src/formats/onenote-file';
import { ImportContext } from '../../src/import-context';
import { parseFrontMatterBlock, serializeFrontMatter } from '../../src/util';
import type { Page, PreservationRecord, Section } from '../../src/formats/onenote-file/semantic/content';
import { MemoryVault, memoryApp } from '../shims/vault';

interface ArchiveAccess {
	ready: Promise<void>;
	indexImportedNotes(): void;
	importSectionArchive(
		ctx: ImportContext,
		section: Section,
		folder: TFolder,
		fallbackName: string,
		groups: string[],
		source: { path: string, sha256: string },
	): Promise<void>;
}

interface PageAccess extends ArchiveAccess {
	planSections(
		prepared: { section: Section, title: string, groups: string[], source: { path: string, sha256: string } }[],
		root: TFolder,
	): {
		item: { section: Section };
		pages: Map<Page, {
			note: { title: string, targetPath: string };
			fullOutputFilename: string;
		}>;
	}[];
	importSection(ctx: ImportContext, planned: unknown, links: Map<string, string[]>): Promise<void>;
}

function section(preservation: PreservationRecord[]): Section {
	return { id: 'section-id', name: 'Section', pages: [], preservation };
}

async function importer() {
	const vault = new MemoryVault();
	const app = memoryApp(vault) as unknown as {
		fileManager: { trashFile(file: TFile): Promise<void> };
	};
	app.fileManager.trashFile = async file => vault.remove(file.path);

	const subject = new OneNoteFileImporter(app as never, {
		sourceEl: null,
		outputEl: null,
		optionsEl: null,
	} as never) as unknown as ArchiveAccess;
	await subject.ready;
	return { vault, subject };
}

async function writeArchive(subject: ArchiveAccess, vault: MemoryVault, preservation: PreservationRecord[]) {
	await subject.importSectionArchive(
		new ImportContext(),
		section(preservation),
		vault.root,
		'Section',
		[],
		{ path: 'C:/Notebook/Section.one', sha256: 'source-hash' });
}

function archiveAsset(vault: MemoryVault) {
	const markdown = vault.contents.get('_OneNote archive.md');
	assert.ok(typeof markdown === 'string');
	const assets = parseFrontMatterBlock(markdown)?.frontMatter['onenote-assets'];
	assert.ok(Array.isArray(assets));
	assert.equal(assets.length, 1);
	return assets[0] as { path: string, length: number, sha256: string };
}

test('the section archive owns every opaque binary through its asset manifest', async () => {
	const { vault, subject } = await importer();
	await writeArchive(subject, vault, [{ code: 'RAW', message: 'opaque', rawData: new Uint8Array([1, 2, 3]) }]);

	const asset = archiveAsset(vault);
	assert.equal(asset.length, 3);
	assert.match(asset.sha256, /^[0-9a-f]{64}$/);
	assert.ok(vault.getAbstractFileByPath(asset.path));
	assert.match(String(vault.contents.get('_OneNote archive.md')), new RegExp(`path="${asset.path.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`));
});

test('reimport removes an unchanged archive binary that the source no longer contains', async () => {
	const { vault, subject } = await importer();
	await writeArchive(subject, vault, [{ code: 'RAW', message: 'opaque', rawData: new Uint8Array([1, 2, 3]) }]);
	const asset = archiveAsset(vault);

	subject.indexImportedNotes();
	await writeArchive(subject, vault, []);

	assert.equal(vault.getAbstractFileByPath(asset.path), null);
});

test('reimport retains stale archive binaries that were modified or are shared', async () => {
	for (const reason of ['modified', 'shared'] as const) {
		const { vault, subject } = await importer();
		await writeArchive(subject, vault, [{ code: 'RAW', message: 'opaque', rawData: new Uint8Array([1, 2, 3]) }]);
		const asset = archiveAsset(vault);
		const file = vault.getAbstractFileByPath(asset.path) as unknown as TFile;

		if (reason === 'modified') await vault.modifyBinary(file, new Uint8Array([9, 2, 3]).buffer);
		else await vault.create('Other.md', serializeFrontMatter({ 'onenote-assets': [asset] }) + 'Still used here');

		subject.indexImportedNotes();
		await writeArchive(subject, vault, []);

		assert.ok(vault.getAbstractFileByPath(asset.path), reason);
	}
});

test('a truncated output filename keeps both complete names in Markdown', async () => {
	const { vault, subject } = await importer();
	const page: Page = {
		id: 'long-page-id',
		title: '完整标题'.repeat(80),
		level: 0,
		isConflictPage: false,
		isDeleted: false,
		outlines: [],
		directContent: [],
		preservation: [],
	};
	const source = { path: 'C:/Notebook/Section.one', sha256: 'source-hash' };
	const input = { section: { ...section([]), pages: [page] }, title: 'Section', groups: [], source };
	const access = subject as PageAccess;
	const [planned] = access.planSections([input], vault.root);

	await access.importSection(new ImportContext(), planned, new Map());

	const pageNote = [...vault.getMarkdownFiles()].find(file => file.path !== '_OneNote archive.md');
	assert.ok(pageNote);
	const markdown = await vault.read(pageNote);
	const frontmatter = parseFrontMatterBlock(markdown)?.frontMatter;
	assert.equal(frontmatter?.title, page.title);
	assert.equal(frontmatter?.['onenote-original-filename'], `${page.title}.md`);
	assert.equal(frontmatter?.['onenote-full-output-filename'], planned.pages.get(page)?.fullOutputFilename);
	assert.equal(frontmatter?.['onenote-output-filename'], pageNote.name);
	assert.notEqual(frontmatter?.['onenote-original-filename'], frontmatter?.['onenote-output-filename']);
});
