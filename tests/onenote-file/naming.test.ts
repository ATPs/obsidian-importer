import { test } from 'node:test';
import assert from 'node:assert/strict';

import { oneNoteLinkNames, oneNotePageNames } from '../../src/formats/onenote-file';
import type { Page } from '../../src/formats/onenote-file/semantic/content';

function page(id: string, title: string, createdUtc?: Date): Page {
	return {
		id,
		title,
		createdUtc,
		level: 0,
		isConflictPage: false,
		isDeleted: false,
		outlines: [],
		directContent: [],
		preservation: [],
	};
}

test('lossy, empty and duplicate page names receive stable timestamp prefixes', () => {
	const created = new Date('2024-11-12T03:04:05Z');
	const pages = [
		page('one', 'A:B', created),
		page('two', 'AB', created),
		page('three', '', created),
		page('four', 'Same', created),
		page('five', 'same', created),
	];
	const names = oneNotePageNames(pages);

	assert.equal(names.get('one'), '20241112-03-04-05 AB');
	assert.equal(names.get('two'), '20241112-03-04-05 AB-1');
	assert.equal(names.get('three'), '20241112-03-04-05 Untitled');
	assert.equal(names.get('four'), '20241112-03-04-05 Same');
	assert.equal(names.get('five'), '20241112-03-04-05 same-1');
	assert.equal(oneNotePageNames(pages).get('one'), names.get('one'), 'names must remain stable between runs');
});

test('the global link index keeps all same-title candidates across sections', () => {
	const pages = [page('one', 'Roadmap'), page('two', 'roadmap'), page('three', 'Unique')];
	const names = oneNotePageNames(pages);
	const links = oneNoteLinkNames(pages, names);

	assert.deepEqual(links.get('roadmap'), [names.get('one'), names.get('two')]);
	assert.deepEqual(links.get('unique'), ['Unique']);
});
