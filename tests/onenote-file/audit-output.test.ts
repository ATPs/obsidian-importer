import { test } from 'node:test';
import assert from 'node:assert/strict';
import * as crypto from 'node:crypto';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { spawnSync } from 'node:child_process';

function writeCandidate(root: string, imageLabel: string, sourceName = 'Direct export figure.png'): void {
	const notebook = path.join(root, 'Notebook');
	const attachments = path.join(notebook, 'attachments');
	fs.mkdirSync(attachments, { recursive: true });
	fs.writeFileSync(path.join(root, '_merge-report.json'), JSON.stringify([{ notebook: 'Notebook', sources: [], sections: [{ pages: 1 }] }]));
	const bytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);
	fs.writeFileSync(path.join(attachments, 'figure.png'), bytes);
	const sha256 = crypto.createHash('sha256').update(bytes).digest('hex');
	fs.writeFileSync(path.join(notebook, 'Page.md'), [
		'---',
		'title: Page',
		'onenote-id: page-id',
		'onenote-assets:',
		'  - path: attachments/figure.png',
		`    length: ${bytes.length}`,
		`    sha256: ${sha256}`,
		`    sourceName: ${JSON.stringify(sourceName)}`,
		'---',
		`![${imageLabel}](attachments/figure.png)`,
	].join('\n'));
}

function audit(root: string): { status: number | null, report: Record<string, unknown> } {
	const result = spawnSync(process.execPath, [
		path.join(process.cwd(), 'node_modules', 'tsx', 'dist', 'cli.mjs'),
		'--tsconfig', 'tsconfig.test.json',
		'scripts/audit-onenote-output.ts', root,
	], { cwd: process.cwd(), encoding: 'utf8' });
	assert.equal(result.error, undefined, result.stderr);
	assert.notEqual(result.stdout, '', result.stderr);
	return { status: result.status, report: JSON.parse(result.stdout) as Record<string, unknown> };
}

test('OneNote output audit accepts a single-line filename image label', () => {
	const root = fs.mkdtempSync(path.join(os.tmpdir(), 'onenote-output-audit-'));
	try {
		writeCandidate(root, 'Direct export figure.png');
		const result = audit(root);
		assert.equal(result.status, 0, JSON.stringify(result.report));
		assert.equal(result.report.failureCount, 0);
		assert.deepEqual((result.report.failures as Record<string, unknown>).imageLabels, []);
	}
	finally {
		fs.rmSync(root, { recursive: true, force: true });
	}
});

test('OneNote output audit rejects a single-line descriptive image label', () => {
	const root = fs.mkdtempSync(path.join(os.tmpdir(), 'onenote-output-audit-'));
	try {
		writeCandidate(root, 'Architecture of chromosome X.png');
		const result = audit(root);
		assert.equal(result.status, 1);
		assert.equal(result.report.failureCount, 1, JSON.stringify(result.report));
		assert.match(JSON.stringify((result.report.failures as Record<string, unknown>).imageLabels), /filename/u);
	}
	finally {
		fs.rmSync(root, { recursive: true, force: true });
	}
});

test('OneNote output audit rejects a multiline OCR image label', () => {
	const root = fs.mkdtempSync(path.join(os.tmpdir(), 'onenote-output-audit-'));
	try {
		writeCandidate(root, 'De novo variants tend to enrich\nCoding\np=O.14');
		const result = audit(root);
		assert.equal(result.status, 1);
		assert.equal(result.report.failureCount, 1, JSON.stringify(result.report));
		assert.match(JSON.stringify((result.report.failures as Record<string, unknown>).imageLabels), /single line/u);
	}
	finally {
		fs.rmSync(root, { recursive: true, force: true });
	}
});
