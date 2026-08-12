import { SvgStroke, strokesToSvg } from '../onenote/ink-svg';
import { Element, Image, Ink, ListInfo, Page, Paragraph, Table, Tag, TextRun } from './semantic/content';
import { PreserveBinary, preservationBlock, preservationXml } from './preservation';

/** Converts half-inch ink units to CSS pixels at 96 DPI. */
const PIXELS_PER_INK_UNIT = 48;

export interface ResolvedAttachment {
	path: string;
	name: string;
	length?: number;
	sha256?: string;
}

export type SkipReason =
	| 'no-data'
	| 'not-representable';

export interface OneNoteConversionOptions {
	/** Writes one asset and answers with the link target, or null to leave it out. */
	saveAttachment: (data: Uint8Array, suggestedName: string) => Promise<ResolvedAttachment | null>;
	/** Turns an internal OneNote page title into the note name written by the importer. */
	/** Answers with every imported page that can be the target. */
	resolveInternalLink?: (pageTitle: string) => string | string[] | undefined;
	onSkipped?: (name: string, reason: SkipReason) => void;
	/**
	 * What the note is called, for the attachments named after it. A page title
	 * can hold characters a file name cannot, so only the importer knows.
	 */
	noteName?: string;
	isCancelled?: () => boolean;
	/** Saves opaque source bytes that Markdown cannot represent. */
	preserveBinary?: PreserveBinary;
}

export interface ConvertedPage {
	markdown: string;
	attachments: ResolvedAttachment[];
	degraded: boolean;
	preservationCount: number;
	cancelled: boolean;
}

const INVISIBLE_MATH = /[\u2061-\u2064]/g;

// Preserve scripts before NFKC folds their glyphs to ordinary characters.
const SUPERSCRIPTS = '⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱ¹²³';
const SUPERSCRIPT_PLAIN = '0123456789+-=()ni123';
const SUBSCRIPTS = '₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎';
const SUBSCRIPT_PLAIN = '0123456789+-=()';

function scriptRuns(text: string, glyphs: string, plain: string, marker: string): string {
	const pattern = new RegExp(`[${glyphs}]+`, 'g');

	return text.replace(pattern, match => {
		const decoded = [...match].map(character => plain[glyphs.indexOf(character)]).join('');
		return `${marker}{${decoded}}`;
	});
}

function toLatex(text: string): string {
	const scripted = scriptRuns(
		scriptRuns(text, SUPERSCRIPTS, SUPERSCRIPT_PLAIN, '^'),
		SUBSCRIPTS, SUBSCRIPT_PLAIN, '_');

	return scripted.normalize('NFKC').replace(INVISIBLE_MATH, '').trim();
}

/** Escapes plain text that Markdown would reinterpret structurally. */
function escapeInline(text: string): string {
	return text.replace(/[[\]`<]/g, '\\$&');
}

function escapeLineStart(line: string): string {
	return line.replace(/^(\s*)(#{1,6}(?=\s|$)|>|\||[-*+](?=\s)|\d+[.)](?=\s)|`{3,}|~{3,}|-{3,}$|={3,}$)/, '$1\\$2');
}

/** Extracts `Page title` from `onenote:...#Page%20title&section-id=...`. */
function internalPageTitle(url: string): string | undefined {
	if (!url.toLowerCase().startsWith('onenote:')) return undefined;

	const hash = url.indexOf('#');
	if (hash < 0) return undefined;

	const tail = url.slice(hash + 1);
	const separator = tail.indexOf('&');
	const encoded = tail.slice(0, separator < 0 ? tail.length : separator);
	if (encoded === '') return undefined;

	try {
		return decodeURIComponent(encoded);
	}
	catch {
		return encoded;
	}
}

function renderRun(run: TextRun, options: OneNoteConversionOptions): string {
	let text = run.text;
	if (text === '') return '';

	const leading = text.match(/^\s*/)![0];
	const trailing = text.length > leading.length ? text.match(/\s*$/)![0] : '';
	let core = text.slice(leading.length, text.length - trailing.length);

	if (core !== '') {
		// Formatting around math would end up inside its delimiters.
		if (run.math) {
			const latex = toLatex(core);
			return latex === '' ? '' : `${leading}$${latex}$${trailing}`;
		}

		core = escapeInline(core);

		if (run.highlight) core = highlighted(core, run.highlight);
		if (run.superscript) core = `<sup>${core}</sup>`;
		if (run.subscript) core = `<sub>${core}</sub>`;
		if (run.underline) core = `<u>${core}</u>`;
		if (run.bold) core = `**${core}**`;
		if (run.italic) core = `*${core}*`;
		if (run.strikethrough) core = `~~${core}~~`;
		if (run.hyperlinkUrl) {
			const pageTitle = internalPageTitle(run.hyperlinkUrl);
			const resolved = pageTitle ? options.resolveInternalLink?.(pageTitle) : undefined;
			const targets = Array.isArray(resolved) ? resolved : resolved === undefined ? [] : [resolved];

			if (!pageTitle || options.resolveInternalLink === undefined) {
				core = `[${core}](${encodeURI(resolved as string | undefined ?? pageTitle ?? run.hyperlinkUrl)})`;
			}
			else if (targets.length === 1) {
				core = `[${core}](${encodeURI(targets[0])})`;
			}
			else if (targets.length > 1) {
				const links = targets.map((target, index) => `[${index + 1}](${encodeURI(target)})`).join(', ');
				core = `${core} *(OneNote link has multiple pages with this title: ${links})*`;
			}
			else {
				core = `[${core}](${encodeURI(run.hyperlinkUrl)}) *(OneNote link target was not found in this import)*`;
			}
		}
	}

	return leading + core + trailing;
}

function runNeedsPreservation(run: TextRun): boolean {
	return run.font !== undefined || run.fontSize !== undefined;
}

function renderRuns(runs: TextRun[], options: OneNoteConversionOptions): string {
	return runs.map(run => renderRun(run, options)).join('').replace(/\r\n?/g, '\n').trim();
}

const HIGHLIGHT_MARKERS: { marker: string, inks: number[][] }[] = [
	{ marker: '🔴', inks: [[0xff, 0x00, 0x00], [0xff, 0x69, 0xb4]] },
	{ marker: '🟠', inks: [[0xff, 0xa5, 0x00]] },
	{ marker: '🟡', inks: [[0xff, 0xff, 0x00]] },
	{ marker: '🟢', inks: [[0x00, 0xff, 0x00], [0x00, 0x80, 0x00]] },
	{ marker: '🔵', inks: [[0x00, 0x00, 0xff], [0x00, 0xff, 0xff]] },
	{ marker: '🟣', inks: [[0x80, 0x00, 0x80], [0xff, 0x00, 0xff]] },
];

function highlighted(text: string, color: string): string {
	const match = color.match(/^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i);
	if (!match) return `==${text}==`;

	const [red, green, blue] = match.slice(1).map(part => parseInt(part, 16));

	let nearest = HIGHLIGHT_MARKERS[0].marker;
	let best = Infinity;

	for (const { marker, inks } of HIGHLIGHT_MARKERS) {
		for (const [inkRed, inkGreen, inkBlue] of inks) {
			const distance = (inkRed - red) ** 2 + (inkGreen - green) ** 2 + (inkBlue - blue) ** 2;
			if (distance < best) {
				best = distance;
				nearest = marker;
			}
		}
	}

	return `==${nearest}${text}==`;
}

function headingPrefix(styleId: string | undefined): string {
	const level = styleId?.match(/^h([1-6])$/i);
	return level ? '#'.repeat(Number(level[1])) + ' ' : '';
}

function listPrefix(list: ListInfo | undefined): string {
	if (!list) return '';
	return '\t'.repeat(list.level) + (list.ordered ? '1. ' : '- ');
}

function taskPrefix(tags: Tag[] | undefined, list: ListInfo | undefined): string | undefined {
	const task = tags?.find(tag => tag.checkable);
	if (!task) return undefined;

	return '\t'.repeat(list?.level ?? 0) + (task.completed ? '- [x] ' : '- [ ] ');
}

// NoteTagShape values that represent admonitions. Labels are localized.
const CALLOUT_SHAPES: Record<number, string> = {
	13: 'important',  // Yellow star
	15: 'question',   // Question mark
	17: 'danger',     // High priority (red exclamation mark)
	21: 'tip',        // Light bulb
	111: 'question',  // Question balloon
};

interface Callout {
	type: string;
	title?: string;
}

function calloutFor(tags: Tag[] | undefined): Callout | undefined {
	for (const tag of tags ?? []) {
		if (tag.checkable || tag.shape === undefined) continue;

		const type = CALLOUT_SHAPES[tag.shape];
		if (type) return { type, title: tag.label };
	}

	return undefined;
}

interface Block {
	text: string;
	listItem: boolean;
	callout?: string;
}

function extensionOf(fileName: string | undefined): string | undefined {
	return fileName?.match(/\.[^.\\/]+$/)?.[0];
}

/** An attachment without an extension is one the vault cannot open. */
function withExtension(base: string, extension: string | undefined): string {
	if (!extension) return base;
	if (extensionOf(base)) return base;
	return base + (extension.startsWith('.') ? extension : `.${extension}`);
}

class PageWriter {
	private readonly blocks: Block[] = [];
	private readonly inkStrokes: SvgStroke[] = [];
	private readonly recognizedText: string[] = [];
	readonly attachments: ResolvedAttachment[] = [];
	readonly preservation: Page['preservation'];
	cancelled = false;

	constructor(private readonly options: OneNoteConversionOptions, private readonly pageTitle: string, initial: Page['preservation']) {
		this.preservation = [...initial];
	}

	get markdown(): string {
		const lines: string[] = [];

		for (const [index, block] of this.blocks.entries()) {
			const previous = this.blocks[index - 1];
			if (previous && !(block.listItem && previous.listItem)) lines.push('');
			lines.push(block.text);
		}

		return lines.join('\n').replace(/\n{3,}/g, '\n\n').trim();
	}

	private push(text: string, listItem = false): void {
		this.blocks.push({ text, listItem });
	}

	private pushCallout(callout: Callout, body: string): void {
		const quoted = body.split('\n').map(line => `> ${line}`).join('\n');
		const previous = this.blocks[this.blocks.length - 1];
		const opening = `> [!${callout.type}]${callout.title ? ` ${callout.title}` : ''}`;

		if (previous?.callout === opening) {
			previous.text += `\n>\n${quoted}`;
			return;
		}

		this.blocks.push({ text: `${opening}\n${quoted}`, listItem: false, callout: opening });
	}

	async writeElements(elements: Element[]): Promise<void> {
		for (const element of elements) {
			if (this.options.isCancelled?.()) {
				this.cancelled = true;
				return;
			}
			await this.writeElement(element);
		}
	}

	private async writeElement(element: Element): Promise<void> {
		switch (element.kind) {
			case 'outline':
				await this.writeElements(element.children);
				break;
			case 'paragraph':
				await this.writeParagraph(element);
				break;
			case 'table':
				await this.writeTable(element);
				break;
			case 'image':
				await this.writeAsset(element.data, this.imageName(element), element.altText ?? '', true);
				break;
			case 'embedded-file': {
				const name = withExtension(element.fileName ?? 'attachment', element.extension);
				await this.writeAsset(element.data, name, name, false);
				break;
			}
			case 'ink':
				this.collectInk(element);
				break;
		}
	}

	private async writeParagraph(paragraph: Paragraph): Promise<void> {
		const text = renderRuns(paragraph.runs, this.options);

		if (text !== '') {
			const task = taskPrefix(paragraph.tags, paragraph.list);
			const prefix = task ?? listPrefix(paragraph.list) ?? '';
			const indent = '\t'.repeat(paragraph.list?.level ?? 0);

			const escaped = text.split('\n').map(escapeLineStart);
			const body = (prefix || headingPrefix(paragraph.styleId)) + escaped.join('  \n' + indent);
			const callout = calloutFor(paragraph.tags);

			if (callout && !paragraph.list && !task) this.pushCallout(callout, body);
			else this.push(body, task !== undefined || paragraph.list !== undefined);
		}

		await this.writeElements(paragraph.children);
	}

	private async writeTable(table: Table): Promise<void> {
		if (table.rows.length === 0) return;

		const columns = Math.max(...table.rows.map(row => row.cells.length));

		const rendered: string[][] = [];
		for (const row of table.rows) {
			const cells: string[] = [];
			for (let index = 0; index < columns; index++) {
				const text = await this.renderCell(row.cells[index]?.children ?? []);
				cells.push(text.replace(/\s+/g, ' ').replace(/\|/g, '\\|').trim());
			}
			rendered.push(cells);
		}

		// GFM has no table without a header, so the first row becomes one.
		const lines = [
			`| ${rendered[0].join(' | ')} |`,
			`| ${new Array(columns).fill('---').join(' | ')} |`,
			...rendered.slice(1).map(row => `| ${row.join(' | ')} |`),
		];

		this.push(lines.join('\n'));
	}

	private collectInk(ink: Ink): void {
		for (const stroke of ink.strokes) {
			this.inkStrokes.push({
				points: stroke.points.map(point => ({ x: point.x * PIXELS_PER_INK_UNIT, y: point.y * PIXELS_PER_INK_UNIT })),
				color: stroke.color,
				width: Math.max(1, stroke.width * PIXELS_PER_INK_UNIT),
				opacity: stroke.opacity,
			});
		}

		// Recognition text is repeated on every stroke in a word.
		if (ink.recognizedText && ink.recognizedText !== this.recognizedText[this.recognizedText.length - 1]) {
			this.recognizedText.push(ink.recognizedText);
		}
	}

	async writeCollectedInk(): Promise<void> {
		const svg = strokesToSvg(this.inkStrokes);
		if (!svg) return;

		const recognized = this.recognizedText.join(' ');
		await this.writeAsset(new TextEncoder().encode(svg), `${this.pageTitle} - Ink.svg`, recognized, true);

		if (recognized !== '') this.push(recognized);
	}

	private imageName(image: Image): string {
		return withExtension(`${this.pageTitle} image`, image.extension ?? extensionOf(image.fileName));
	}

	private async writeAsset(data: Uint8Array | undefined, name: string, label: string, embed: boolean): Promise<void> {
		const link = await this.renderAsset(data, name, label, embed);
		if (link) this.push(link);
	}

	private async renderAsset(data: Uint8Array | undefined, name: string, label: string, embed: boolean): Promise<string | undefined> {
		if (!data || data.length === 0) {
			this.options.onSkipped?.(name, 'no-data');
			this.preservation.push({
				code: 'ONENOTE_ASSET_DATA_MISSING',
				message: 'OneNote described an attachment or image but its payload was not available.',
				details: { name, label, embed },
			});
			return undefined;
		}

		const attachment = await this.options.saveAttachment(data, name);
		if (!attachment) {
			this.options.onSkipped?.(name, 'no-data');
			this.preservation.push({
				code: 'ONENOTE_ASSET_NOT_SAVED',
				message: 'An attachment or image was read but the destination refused it.',
				details: { name, label, embed, length: data.length },
				rawData: data,
			});
			return undefined;
		}

		this.attachments.push(attachment);
		const target = encodeURI(attachment.path);
		return embed ? `![${label}](${target})` : `[${label}](${target})`;
	}

	private async renderCell(children: Element[]): Promise<string> {
		const parts: string[] = [];

		for (const child of children) {
			switch (child.kind) {
				case 'paragraph':
					parts.push(renderRuns(child.runs, this.options));
					parts.push(await this.renderCell(child.children));
					break;
				case 'outline':
					parts.push(await this.renderCell(child.children));
					break;
				case 'image':
					parts.push(await this.renderAsset(child.data, this.imageName(child), child.altText ?? '', true) ?? '');
					break;
				case 'embedded-file': {
					const name = withExtension(child.fileName ?? 'attachment', child.extension);
					parts.push(await this.renderAsset(child.data, name, name, false) ?? '');
					break;
				}
				case 'ink':
					this.collectInk(child);
					break;
				case 'table':
					this.options.onSkipped?.(this.pageTitle, 'not-representable');
					await this.preserveNestedTableAssets(child);
					this.preservation.push({
						code: 'ONENOTE_NESTED_TABLE',
						message: 'A table nested inside a table cell has no equivalent in Markdown tables.',
						details: { rows: child.rows.length },
					});
					break;
			}
		}

		return parts.filter(part => part !== '').join(' ');
	}

	private async preserveNestedTableAssets(table: Table): Promise<void> {
		const visit = async (element: Element): Promise<void> => {
			if (element.kind === 'image' || element.kind === 'embedded-file') {
				if (!element.data || element.data.length === 0) return;
				const name = element.kind === 'image'
					? withExtension(element.fileName ?? `${this.pageTitle} nested-table image`, element.extension)
					: withExtension(element.fileName ?? 'nested-table attachment', element.extension);
				this.preservation.push({
					code: 'ONENOTE_NESTED_TABLE_ASSET',
					message: 'An asset inside an unrepresentable nested table is preserved as opaque source data.',
					details: { name, kind: element.kind },
					rawData: element.data,
				});
				return;
			}
			if (element.kind === 'paragraph' || element.kind === 'outline') {
				for (const child of element.children) await visit(child);
			}
			else if (element.kind === 'table') {
				for (const row of element.rows) for (const cell of row.cells) for (const child of cell.children) await visit(child);
			}
		};

		await visit(table);
	}
}

export async function convertPage(page: Page, options: OneNoteConversionOptions): Promise<ConvertedPage> {
	const writer = new PageWriter(options, options.noteName ?? page.title, page.preservation);

	await writer.writeElements(page.outlines);
	await writer.writeElements(page.directContent);
	await writer.writeCollectedInk();

	const records = writer.preservation;
	const formattedRuns: {
		font?: string;
		fontSize?: number;
		highlight?: string;
		hyperlinkUrl?: string;
		math?: boolean;
		text: string;
	}[] = [];
	const taggedParagraphs: { text: string, tags: Tag[] }[] = [];
	const specialAssets: Record<string, string | number>[] = [];
	const visit = (element: Element): void => {
		if (element.kind === 'paragraph') {
			for (const run of element.runs) {
				if (runNeedsPreservation(run) || run.highlight || run.hyperlinkUrl || run.math) {
					formattedRuns.push({
						text: run.text,
						font: run.font,
						fontSize: run.fontSize,
						highlight: run.highlight,
						hyperlinkUrl: run.hyperlinkUrl,
						math: run.math,
					});
				}
			}
			if (element.tags?.length) taggedParagraphs.push({ text: element.runs.map(run => run.text).join(''), tags: element.tags });
			for (const child of element.children) visit(child);
		}
		else if (element.kind === 'outline') element.children.forEach(visit);
		else if (element.kind === 'table') {
			for (const row of element.rows) for (const cell of row.cells) cell.children.forEach(visit);
		}
		else if (element.kind === 'image') {
			specialAssets.push({ kind: 'image', originalName: element.fileName ?? '', altText: element.altText ?? '', extension: element.extension ?? '' });
		}
		else if (element.kind === 'embedded-file') {
			specialAssets.push({
				kind: 'embedded-file',
				originalName: element.fileName ?? '',
				sourcePath: element.sourcePath ?? '',
				extension: element.extension ?? '',
			});
		}
	};
	page.outlines.forEach(visit);
	page.directContent.forEach(visit);

	if (formattedRuns.length > 0) {
		const details: Record<string, string | number | boolean> = {};
		for (const [index, run] of formattedRuns.entries()) {
			details[`run-${index}-text`] = run.text;
			if (run.font !== undefined) details[`run-${index}-font`] = run.font;
			if (run.fontSize !== undefined) details[`run-${index}-font-size`] = run.fontSize;
			if (run.highlight !== undefined) details[`run-${index}-highlight`] = run.highlight;
			if (run.hyperlinkUrl !== undefined) details[`run-${index}-hyperlink`] = run.hyperlinkUrl;
			if (run.math !== undefined) details[`run-${index}-math`] = run.math;
		}
		records.push({
			code: 'ONENOTE_TEXT_SOURCE_FORMAT',
			message: 'OneNote text formatting or source link cannot be represented reversibly in portable Markdown.',
			details,
		});
	}
	for (const [index, paragraph] of taggedParagraphs.entries()) {
		const details: Record<string, string | number | boolean> = { text: paragraph.text };
		for (const [tagIndex, tag] of paragraph.tags.entries()) {
			if (tag.label !== undefined) details[`tag-${tagIndex}-label`] = tag.label;
			if (tag.shape !== undefined) details[`tag-${tagIndex}-shape`] = tag.shape;
			details[`tag-${tagIndex}-checkable`] = tag.checkable;
			details[`tag-${tagIndex}-completed`] = tag.completed;
		}
		records.push({
			code: 'ONENOTE_TAG_SOURCE',
			message: 'The exact OneNote tag definition and state are preserved alongside their Markdown approximation.',
			details: { paragraph: index, ...details },
		});
	}
	for (const [index, details] of specialAssets.entries()) {
		records.push({
			code: 'ONENOTE_ASSET_SOURCE_METADATA',
			message: 'Original OneNote attachment metadata is preserved because the vault path may be sanitized or renamed.',
			details: { asset: index, ...details },
		});
	}

	let markdown = writer.markdown;
	if (records.length > 0 && options.preserveBinary) {
		const xml = await preservationXml(records, options.preserveBinary);
		markdown = [markdown, preservationBlock(xml)].filter(part => part !== '').join('\n\n');
	}

	return {
		markdown,
		attachments: writer.attachments,
		degraded: records.length > 0,
		preservationCount: records.length,
		cancelled: writer.cancelled,
	};
}
