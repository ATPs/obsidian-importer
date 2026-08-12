import { test } from 'node:test';
import assert from 'node:assert/strict';

import { oneNoteLinkNames, oneNotePageNames, shortOneNoteId } from '../../src/formats/onenote-file';
import type { Page } from '../../src/formats/onenote-file/semantic/content';

function page(id: string, title: string): Page {
	return {
		id,
		title,
		level: 0,
		isConflictPage: false,
		isDeleted: false,
		outlines: [],
		directContent: [],
		preservation: [],
	};
}

test('lossy, empty and duplicate page names receive stable short identifiers', () => {
	const pages = [
		page('one', 'A:B'),
		page('two', 'AB'),
		page('three', ''),
		page('four', 'Same'),
		page('five', 'same'),
	];
	const names = oneNotePageNames(pages);

	assert.equal(names.get('one'), `${shortOneNoteId('one')} AB`);
	assert.equal(names.get('two'), `${shortOneNoteId('two')} AB`);
	assert.equal(names.get('three'), `${shortOneNoteId('three')} Untitled`);
	assert.equal(names.get('four'), `${shortOneNoteId('four')} Same`);
	assert.equal(names.get('five'), `${shortOneNoteId('five')} same`);
	assert.equal(oneNotePageNames(pages).get('one'), names.get('one'), 'names must remain stable between runs');
});

test('the global link index keeps all same-title candidates across sections', () => {
	const pages = [page('one', 'Roadmap'), page('two', 'roadmap'), page('three', 'Unique')];
	const names = oneNotePageNames(pages);
	const links = oneNoteLinkNames(pages, names);

	assert.deepEqual(links.get('roadmap'), [names.get('one'), names.get('two')]);
	assert.deepEqual(links.get('unique'), ['Unique']);
});
