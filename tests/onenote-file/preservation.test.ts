import { test } from 'node:test';
import assert from 'node:assert/strict';

import { preservationBlock, preservationXml, sha256Hex } from '../../src/formats/onenote-file/preservation';

test('preservation XML escapes source text and retains exact binary bytes', async () => {
	const saved: { name: string, data: Uint8Array }[] = [];
	const document = await preservationXml([{
		code: 'UNKNOWN<&',
		message: 'cannot <render> & keep "all" bytes',
		objectId: 'object:1',
		jcid: 0x600ff,
		properties: [{ rawId: 0x1c001234, index: 0, data: new Uint8Array([0, 1, 255]) }],
	}], async (data, name) => {
		saved.push({ name, data });
		return { path: `attachments/${name}`, length: data.length, sha256: await sha256Hex(data) };
	});

	assert.match(document, /code="UNKNOWN&lt;&amp;"/);
	assert.match(document, /cannot &lt;render&gt; &amp; keep &quot;all&quot; bytes/);
	assert.match(document, /<binary encoding="base64" length="3">AAH\/<\/binary>/);
	assert.deepEqual(saved, []);
});

test('small exact property bytes stay inline instead of creating thousands of sidecars', async () => {
	let sidecars = 0;
	const document = await preservationXml([{
		code: 'PROPERTY',
		message: 'exact bytes',
		properties: [{ rawId: 1, index: 0, data: new Uint8Array([0, 1, 2, 253, 254, 255]) }],
	}], async data => {
		sidecars++;
		return { path: 'unexpected.bin', length: data.length, sha256: 'hash' };
	});

	assert.equal(sidecars, 0);
	assert.match(document, /<binary encoding="base64" length="6">AAEC\/f7\/</);
});

test('large property bytes remain a hashed sidecar', async () => {
	const data = new Uint8Array(64 * 1024 + 1);
	data[0] = 1;
	data[data.length - 1] = 2;
	const saved: { name: string, data: Uint8Array }[] = [];
	const document = await preservationXml([{
		code: 'LARGE_PROPERTY',
		message: 'exact bytes',
		properties: [{ rawId: 1, index: 4, data }],
	}], async (bytes, name) => {
		saved.push({ name, data: bytes });
		return { path: `attachments/${name}`, length: bytes.length, sha256: await sha256Hex(bytes) };
	});

	assert.equal(saved.length, 1);
	assert.equal(saved[0].name, 'preserved-1-property-4.bin');
	assert.equal(saved[0].data, data);
	assert.match(document, /<binary path="attachments\/preserved-1-property-4\.bin" length="65537" sha256="[0-9a-f]{64}"\/>/);
});

test('preservation data uses an ordinary Markdown XML fence', () => {
	assert.equal(
		preservationBlock('<onenote-preservation/>'),
		'## OneNote preservation data\n\n```xml\n<onenote-preservation/>\n```');
});
