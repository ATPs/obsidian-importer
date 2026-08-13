import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

import { convertOneNoteLocal } from '../../scripts/convert-onenote-local';
import { exportOneNoteBackups, linkTargetForAsset, readMissingAssets, rewriteFinalLinks, sameAttachmentLabel, writeMissingAssets } from '../../scripts/export-onenote-backups';

const fixture = path.join(__dirname, 'fixtures', 'testOneNote.one');

test('local OneNote export puts attachments beside the note in attachments', async () => {
	const root = fs.mkdtempSync(path.join(os.tmpdir(), 'onenote-local-export-'));
	const output = path.join(root, 'output');

	try {
		await convertOneNoteLocal(fixture, output);
		const note = path.join(output, 'testOneNote', 'Note-ssn-test-mmmm.md');
		const attachment = path.join(output, 'testOneNote', 'attachments', 'Note-ssn-test-mmmm image.png');
		assert.ok(fs.existsSync(note));
		assert.ok(fs.existsSync(attachment));
		assert.match(fs.readFileSync(note, 'utf8'), /!\[.*\]\(testOneNote\/attachments\/Note-ssn-test-mmmm%20image\.png\)/);
	}
	finally {
		fs.rmSync(root, { recursive: true, force: true });
	}
});

test('backup merge reads structured missing assets and ignores malformed entries', () => {
	assert.deepEqual(readMissingAssets({
		'onenote-missing-assets': [
			{ name: 'diagram.emf', label: 'diagram', embed: true },
			{ name: 'missing-label.bin', embed: false },
		],
	}), [{ name: 'diagram.emf', label: 'diagram', embed: true }]);
});

test('backup merge matches an OCR attachment label despite generated code fences', () => {
	const current = 'De novo variants tend to enrich\nCoding\np=O.14\nRR=I .57\nLGD\nTD Male';
	const older = 'De novo variants tend to enrich\nCoding\n```python\np=O.14\nRR=I .57\n```\nLGD\nTD Male';
	assert.equal(sameAttachmentLabel(current, older), true);
	assert.equal(sameAttachmentLabel(current, 'Architecture of Chromosome X\nPAR1\nPAR2\nCentromere'), false);
});

test('backup merge removes the missing asset field after recovery and keeps unrecovered entries', () => {
	const frontMatter: Record<string, unknown> = { 'onenote-missing-assets': [{ name: 'old.bin', label: 'old', embed: false }] };
	writeMissingAssets(frontMatter, []);
	assert.equal('onenote-missing-assets' in frontMatter, false);

	const remaining = [{ name: 'still-missing.bin', label: 'missing', embed: false }];
	writeMissingAssets(frontMatter, remaining);
	assert.deepEqual(frontMatter['onenote-missing-assets'], remaining);
});

test('backup merge reads an attachment label containing an unescaped bracket', () => {
	const asset = { path: 'Section/attachments/page image 1.png', length: 1, sha256: 'hash' };
	const label = 'Isoform expression [source:HGNC\nENST000001';
	const body = `![${label}](${encodeURI(asset.path)})`;

	assert.deepEqual(linkTargetForAsset(body, asset), { label, embed: true });
});

test('backup merge rewrites unique staging note links to final relative paths', () => {
	const index = new Map([
		['section.one (于 old)/parent/target', ['Section/Parent/Target.md']],
	]);
	const body = [
		'[target](Section.one%20(%E4%BA%8E%20old)/parent/target "jump")',
		'[external](https://example.com/Section.one)',
		'[unresolved](Section.one%20(%E4%BA%8E%20old)/missing)',
	].join('\n');

	assert.equal(rewriteFinalLinks(body, 'Section.one (于 old)/parent/source', 'Section/Parent/Source.md', index), [
		'[target](./Target "jump")',
		'[external](https://example.com/Section.one)',
		'[unresolved](Section.one%20(%E4%BA%8E%20old)/missing)',
	].join('\n'));
});

test('backup merge rewrites OneMore links with escaped brackets in their labels', () => {
	const index = new Map([
		['section.one (于 old)/target', ['Section/Target.md']],
	]);
	const body = '[# heading \\[Top of page\\]](Section.one%20(%E4%BA%8E%20old)/target)';

	assert.equal(
		rewriteFinalLinks(body, 'Section.one (于 old)/source', 'Section/Source.md', index),
		'[# heading \\[Top of page\\]](./Target)');
});

test('backup merge rewrites a link surrounded by prose parentheses', () => {
	const index = new Map([
		['postcomet.one (于 3-21-2024 - 2)/0928 summary statistics highlights', ['PostComet/0928 summary statistics highlights.md']],
	]);
	const body = 'peptides table ([target](PostComet.one%20(%E4%BA%8E%203-21-2024%20-%202)/0928%20summary%20statistics%20highlights))';

	assert.equal(
		rewriteFinalLinks(body, 'PostComet.one (于 3-21-2024 - 2)/parent/source', 'PostComet/Parent/Source.md', index),
		'peptides table ([target](../0928%20summary%20statistics%20highlights))');
});

test('backup merge leaves ambiguous staging note links unchanged', () => {
	const destination = 'Section.one%20(%E4%BA%8E%20old)/duplicate';
	const index = new Map([
		['section.one (于 old)/duplicate', ['Section/Duplicate.md', 'Section/Other Duplicate.md']],
	]);
	assert.equal(rewriteFinalLinks(`[duplicate](${destination})`, 'Section.one (于 old)/source', 'Section/Source.md', index), `[duplicate](${destination})`);
});

test('backup merge rewrites a double-encoded OneNote link by unique final title', () => {
	const destination = 'onenote:Section.one#Other%2520page&section-id=%7Bsection%7D&page-id=%7Bpage%7D';
	const body = `[other](${destination}) *(OneNote link target was not found in this import)*`;
	const index = new Map([['title:other page', ['Other Section/Other page.md']]]);

	assert.equal(
		rewriteFinalLinks(body, 'Section/Source', 'Section/Source.md', index),
		'[other](../Other%20Section/Other%20page)');
});

test('backup merge keeps only visible text for an ambiguous OneNote title', () => {
	const destination = 'onenote:Section.one#Duplicate&section-id=%7Bsection%7D&page-id=%7Bpage%7D';
	const body = `[duplicate](${destination}) *(OneNote link target was not found in this import)*`;
	const index = new Map([['title:duplicate', ['A/Duplicate.md', 'B/Duplicate.md']]]);

	assert.equal(rewriteFinalLinks(body, 'Section/Source', 'Section/Source.md', index), 'duplicate');
});

test('backup merge keeps only visible text for a missing OneNote target', () => {
	const destination = 'onenote:Section.one#Missing%2520page&section-id=%7Bsection%7D&page-id=%7Bpage%7D';
	const body = `[shown title](${destination}) *(OneNote link target was not found in this import)*`;

	assert.equal(rewriteFinalLinks(body, 'Section/Source', 'Section/Source.md', new Map()), 'shown title');
});

test('backup merge converts a leaked OneNote HYPERLINK field code', () => {
	const body = '\uf7dfHYPERLINK "onenote:#Other%2520page&section-id={section}"visible text';
	const index = new Map([['title:other page', ['Other Section/Other page.md']]]);

	assert.equal(
		rewriteFinalLinks(body, 'Section/Source', 'Section/Source.md', index),
		'[visible text](../Other%20Section/Other%20page)');
});

test('global backup merge keeps only the title for a cross-notebook OneNote link', () => {
	const destination = 'onenote:Section.one#Other%2520page&section-id=%7Bsection%7D';
	const body = `[other page](${destination}) *(OneNote link target was not found in this import)*`;
	const index = new Map([['title:other page', ['Notebook B/Section/Other page.md']]]);

	assert.equal(
		rewriteFinalLinks(body, 'Notebook A/Section/Source', 'Notebook A/Section/Source.md', index, 'title'),
		'other page');
});

test('backup merge does not publish the OneNote archive note', async () => {
	const root = fs.mkdtempSync(path.join(os.tmpdir(), 'onenote-backup-export-'));
	const input = path.join(root, 'Notebook');
	const output = path.join(root, 'output');
	fs.mkdirSync(input);
	fs.copyFileSync(fixture, path.join(input, 'Section.one'));

	try {
		await exportOneNoteBackups(output, [input]);
		const archives: string[] = [];
		const walk = (directory: string): void => {
			for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
				const item = path.join(directory, entry.name);
				if (entry.isDirectory()) walk(item);
				else if (entry.name.toLowerCase() === '_onenote archive.md') archives.push(item);
			}
		};
		walk(path.join(output, 'Notebook'));
		assert.deepEqual(archives, []);
	}
	finally {
		fs.rmSync(root, { recursive: true, force: true });
	}
});
