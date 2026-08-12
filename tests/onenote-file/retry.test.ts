import { test } from 'node:test';
import assert from 'node:assert/strict';

import { OneNoteFormatError } from '../../src/formats/onenote-file/errors';
import { isTransientOneNoteError, retryTransient } from '../../src/formats/onenote-file/retry';

test('temporary Windows file errors retry with exponential backoff', async () => {
	let calls = 0;
	const delays: number[] = [];
	const retries: number[] = [];

	const result = await retryTransient(async () => {
		calls++;
		if (calls < 3) throw Object.assign(new Error('sharing violation'), { code: 'EBUSY' });
		return 'written';
	}, {
		baseDelayMs: 10,
		wait: async delay => { delays.push(delay); },
		onRetry: attempt => retries.push(attempt),
	});

	assert.equal(result, 'written');
	assert.equal(calls, 3);
	assert.deepEqual(delays, [10, 20]);
	assert.deepEqual(retries, [2, 3]);
});

test('deterministic OneNote format errors are not retried', async () => {
	let calls = 0;
	await assert.rejects(
		retryTransient(async () => {
			calls++;
			throw new OneNoteFormatError('ONENOTE_UNKNOWN_FILE_FORMAT', 'unsupported');
		}, { wait: async () => assert.fail('format errors must not wait') }),
		OneNoteFormatError);
	assert.equal(calls, 1);
});

test('the retry classifier recognises transient failures but not ordinary bugs', () => {
	assert.equal(isTransientOneNoteError(Object.assign(new Error('locked'), { code: 'EPERM' })), true);
	assert.equal(isTransientOneNoteError(new Error('Cannot read property x')), false);
});
