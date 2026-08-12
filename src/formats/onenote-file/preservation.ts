import { PreservedProperty, PreservationRecord } from './semantic/content';
import { nodeCrypto } from '../../filesystem';

export interface PreservedBinary {
	path: string;
	length: number;
	sha256: string;
}

export type PreserveBinary = (data: Uint8Array, suggestedName: string) => Promise<PreservedBinary>;

const XML_NAME = /^[A-Za-z_][A-Za-z0-9_.-]*$/;
const INLINE_PROPERTY_BYTES = 64 * 1024;

function xml(value: unknown): string {
	return String(value)
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&apos;');
}

function hex(value: number): string {
	return `0x${value.toString(16).padStart(8, '0')}`;
}

function base64(data: Uint8Array): string {
	let encoded = '';
	const chunkSize = 0x8000;
	for (let offset = 0; offset < data.length; offset += chunkSize) {
		encoded += String.fromCharCode(...data.subarray(offset, offset + chunkSize));
	}
	return btoa(encoded);
}

function attributes(values: object): string {
	return Object.entries(values)
		.filter(([, value]) => value !== undefined)
		.map(([name, value]) => ` ${XML_NAME.test(name) ? name : 'value'}="${xml(value)}"`)
		.join('');
}

export async function sha256Hex(data: Uint8Array): Promise<string> {
	if (nodeCrypto) return nodeCrypto.createHash('sha256').update(data).digest('hex');

	const copy = new Uint8Array(data.length);
	copy.set(data);
	const digest = await crypto.subtle.digest('SHA-256', copy.buffer);
	return [...new Uint8Array(digest)].map(byte => byte.toString(16).padStart(2, '0')).join('');
}

async function propertyXml(
	property: PreservedProperty,
	prefix: string,
	preserveBinary: PreserveBinary,
): Promise<string[]> {
	const lines = [`<property${attributes({ id: hex(property.rawId & 0x03ffffff), rawId: hex(property.rawId), index: property.index })}>`];

	if (property.booleanValue !== undefined) lines.push(`  <boolean>${property.booleanValue}</boolean>`);
	if (property.scalarValue !== undefined) lines.push(`  <scalar>${property.scalarValue}</scalar>`);
	if (property.referencedIds) {
		for (const id of property.referencedIds) lines.push(`  <reference id="${xml(id)}"/>`);
	}
	if (property.data && property.data.length > 0) {
		if (property.data.length <= INLINE_PROPERTY_BYTES) {
			lines.push(`  <binary encoding="base64" length="${property.data.length}">${base64(property.data)}</binary>`);
		}
		else {
			const saved = await preserveBinary(property.data, `${prefix}-property-${property.index}.bin`);
			lines.push(`  <binary${attributes(saved)}/>`);
		}
	}
	if (property.children) {
		for (const [setIndex, set] of property.children.entries()) {
			lines.push(`  <property-set${attributes({ index: setIndex, childPropertyId: property.childPropertyId === undefined ? undefined : hex(property.childPropertyId) })}>`);
			for (const child of set) {
				lines.push(...(await propertyXml(child, `${prefix}-${property.index}-${setIndex}`, preserveBinary)).map(line => `    ${line}`));
			}
			lines.push('  </property-set>');
		}
	}

	lines.push('</property>');
	return lines;
}

export async function preservationXml(
	records: PreservationRecord[],
	preserveBinary: PreserveBinary,
): Promise<string> {
	const lines = ['<onenote-preservation version="1">'];

	for (const [index, record] of records.entries()) {
		const prefix = `preserved-${index + 1}`;
		lines.push(`  <record${attributes({
			code: record.code,
			objectId: record.objectId,
			jcid: record.jcid === undefined ? undefined : hex(record.jcid),
			offset: record.offset === undefined ? undefined : hex(record.offset),
		})}>`);
		lines.push(`    <message>${xml(record.message)}</message>`);

		for (const [name, value] of Object.entries(record.details ?? {})) {
			lines.push(`    <detail name="${xml(name)}" value="${xml(value)}"/>`);
		}

		if (record.rawData && record.rawData.length > 0) {
			const saved = await preserveBinary(record.rawData, `${prefix}-raw.bin`);
			lines.push(`    <raw-binary${attributes(saved)}/>`);
		}

		for (const property of record.properties ?? []) {
			lines.push(...(await propertyXml(property, prefix, preserveBinary)).map(line => `    ${line}`));
		}

		lines.push('  </record>');
	}

	lines.push('</onenote-preservation>');
	return lines.join('\n');
}

export function preservationBlock(xmlDocument: string): string {
	return ['## OneNote preservation data', '', '```xml', xmlDocument, '```'].join('\n');
}
