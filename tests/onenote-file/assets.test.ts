import '../shims/runtime';

import { test } from 'node:test';
import assert from 'node:assert/strict';
import type { TFolder } from 'obsidian';

import { OneNoteFileImporter } from '../../src/formats/onenote-file';
import { ImportContext } from '../../src/import-context';
import { parseFrontMatterBlock } from '../../src/util';
import type { Page, Section } from '../../src/formats/onenote-file/semantic/content';
import { MemoryVault, memoryApp } from '../shims/vault';

interface PageAccess {
	ready: Promise<void>;
	indexImportedNotes(): void;
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

function pageWithAssets(...data: Uint8Array[]): Page {
	return {
		id: 'page-id',
		title: 'Page',
		level: 0,
		isConflictPage: false,
		isDeleted: false,
		outlines: [],
		directContent: data.map(bytes => ({ kind: 'image', extension: '.png', data: bytes })),
		preservation: [],
	};
}

function section(): Section {
	return { id: 'section-id', name: 'Section', pages: [], preservation: [] };
}

async function importer() {
	const vault = new MemoryVault();
	const subject = new OneNoteFileImporter(memoryApp(vault), {
		sourceEl: null,
		outputEl: null,
		optionsEl: null,
	} as never) as unknown as PageAccess;
	await subject.ready;
	return { vault, subject };
}

async function writePages(subject: PageAccess, vault: MemoryVault, ctx: ImportContext, pages: Page[]) {
	const source = { path: 'C:/Notebook/Section.one', sha256: 'source-hash' };
	const input = { section: { ...section(), pages }, title: 'Section', groups: [], source };
	const [planned] = subject.planSections([input], vault.root);
	await subject.importSection(ctx, planned, new Map());
}

test('page note write failure removes a newly created asset', async () => {
	const { vault, subject } = await importer();
	const create = vault.create.bind(vault);
	vault.create = async (path, data, options) => {
		if (path.endsWith('/Page.md')) throw new Error('injected page write failure');
		return await create(path, data, options);
	};
	const ctx = new ImportContext();

	await writePages(subject, vault, ctx, [pageWithAssets(new Uint8Array([1, 2, 3]))]);

	assert.equal(ctx.attachments, 0);
	assert.equal(vault.getAbstractFileByPath('Page image.png'), null);
	assert.ok(ctx.failed.includes('Page'));
});

test('page note write failure retains a reused asset', async () => {
	const { vault, subject } = await importer();
	const bytes = new Uint8Array([1, 2, 3]);
	await vault.createBinary('Page image.png', bytes.buffer);
	const create = vault.create.bind(vault);
	vault.create = async (path, data, options) => {
		if (path.endsWith('/Page.md')) throw new Error('injected page write failure');
		return await create(path, data, options);
	};
	const ctx = new ImportContext();

	await writePages(subject, vault, ctx, [pageWithAssets(bytes)]);

	assert.equal(ctx.attachments, 0);
	assert.ok(vault.getAbstractFileByPath('Page image.png'));
	assert.deepEqual(new Uint8Array(await vault.readBinary({ path: 'Page image.png' })), bytes);
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
	const input = { section: { ...section(), pages: [page] }, title: 'Section', groups: [], source };
	const access = subject;
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
