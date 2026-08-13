import { Resvg } from '@resvg/resvg-js';

const EMF_RENDER_DPI = 300;
const MAX_SHORT_EDGE = 2000;

export interface EmfPng {
	bytes: Uint8Array;
	width: number;
	height: number;
}

function normalizedSvg(svg: string): string {
	// emf-to-png's bundled converter emits svg:defs/clipPath/path but closes
	// some of them without the prefix. Resvg rejects that otherwise valid
	// drawing stream. The namespace prefix carries no semantic distinction.
	return svg.replace(/<\/?svg:/gu, match => match.replace('svg:', ''));
}

function pngDimensions(bytes: Uint8Array): { width: number, height: number } {
	if (bytes.length < 24
		|| bytes[0] !== 0x89 || bytes[1] !== 0x50 || bytes[2] !== 0x4e || bytes[3] !== 0x47
		|| bytes[4] !== 0x0d || bytes[5] !== 0x0a || bytes[6] !== 0x1a || bytes[7] !== 0x0a) {
		throw new Error('EMF conversion did not produce a PNG file.');
	}

	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	const width = view.getUint32(16);
	const height = view.getUint32(20);
	if (width === 0 || height === 0) throw new Error('EMF conversion produced an empty PNG canvas.');
	return { width, height };
}

/** Renders an EMF at 300 DPI, reducing only oversized canvases by short edge. */
export async function convertEmfToPng(bytes: Uint8Array): Promise<EmfPng> {
	// This package only exposes an ESM import entry. A dynamic import keeps the
	// local tsx script compatible with its CommonJS test runner as well.
	const { emfOrWmfToSvg } = await import('emf-to-png');
	const svg = normalizedSvg(await emfOrWmfToSvg('emf', bytes, EMF_RENDER_DPI));
	let rendered = new Uint8Array(new Resvg(svg, { background: '#ffffff' }).render().asPng());
	let { width, height } = pngDimensions(rendered);
	const shortEdge = Math.min(width, height);

	if (shortEdge > MAX_SHORT_EDGE) {
		const scale = MAX_SHORT_EDGE / shortEdge;
		const targetWidth = Math.max(1, Math.round(width * scale));
		rendered = new Uint8Array(new Resvg(svg, {
			background: '#ffffff',
			fitTo: { mode: 'width', value: targetWidth },
		}).render().asPng());
		({ width, height } = pngDimensions(rendered));
		if (Math.min(width, height) > MAX_SHORT_EDGE) {
			throw new Error(`EMF conversion produced an oversized ${width}x${height} PNG.`);
		}
	}

	return { bytes: rendered, width, height };
}

export const emfPngSettings = { dpi: EMF_RENDER_DPI, maxShortEdge: MAX_SHORT_EDGE };
