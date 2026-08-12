import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

import { convertOneNoteLocal } from '../../scripts/convert-onenote-local';

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
