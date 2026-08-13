import { SvgStroke, strokesToSvg } from '../onenote/ink-svg';
import { Element, Image, Ink, ListInfo, Page, Paragraph, Table, Tag, TextRun } from './semantic/content';

/** Converts half-inch ink units to CSS pixels at 96 DPI. */
const PIXELS_PER_INK_UNIT = 48;

export interface ResolvedAttachment {
	path: string;
	name: string;
	length?: number;
	sha256?: string;
	sourceName?: string;
	ordinal?: number;
	embed?: boolean;
}

export interface AssetSource {
	sourceName?: string;
	ordinal?: number;
	embed: boolean;
}

export type SkipReason =
	| 'no-data'
	| 'not-representable';

export interface OneNoteConversionOptions {
	/** Writes one asset and answers with the link target, or null to leave it out. */
	saveAttachment: (data: Uint8Array, suggestedName: string, source?: AssetSource) => Promise<ResolvedAttachment | null>;
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
}

export interface MissingAsset {
	name: string;
	label: string;
	embed: boolean;
	sourceName?: string;
	ordinal?: number;
}

export interface ConvertedPage {
	markdown: string;
	attachments: ResolvedAttachment[];
	missingAssets: MissingAsset[];
	htmlFallbacks: number;
	degraded: boolean;
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
		let decoded = decodeURIComponent(encoded);
		if (/%[0-9a-f]{2}/i.test(decoded)) decoded = decodeURIComponent(decoded);
		return decoded;
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
			const scheme = run.hyperlinkUrl.toLowerCase();
			const isOneNoteLink = scheme.startsWith('onenote:');
			const isOneMoreLink = scheme.startsWith('onemore:');

			if (!isOneNoteLink && !isOneMoreLink) core = `[${core}](${encodeURI(run.hyperlinkUrl)})`;
			else if (targets.length === 1) {
				core = `[${core}](${encodeURI(targets[0])})`;
			}
			// A cross-notebook, missing, or ambiguous target cannot be mapped safely.
			// Keep the visible text without retaining a OneNote-specific URL.
		}
	}

	return leading + core + trailing;
}

function renderRuns(runs: TextRun[], options: OneNoteConversionOptions): string {
	return runs.map(run => renderRun(run, options)).join('').replace(/\r\n?|\v/g, '\n').trim();
}

const MONOSPACE_FONTS = new Set([
	'cascadia code',
	'cascadia mono',
	'consolas',
	'courier',
	'courier new',
	'lucida console',
	'menlo',
	'monaco',
	'source code pro',
]);

function rawParagraphText(paragraph: Paragraph): string {
	return paragraph.runs.map(run => run.text).join('').replace(/\r\n?|\v/g, '\n');
}

function redHighlightedParagraph(paragraph: Paragraph): boolean {
	const visible = paragraph.runs.filter(run => run.text.trim() !== '');
	return visible.length > 0 && visible.every(run => {
		const match = run.highlight?.match(/^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i);
		if (!match) return false;
		const [red, green, blue] = match.slice(1).map(part => parseInt(part, 16));
		return red >= 180 && red >= green * 1.5 && red >= blue * 1.5;
	});
}

function simpleParagraph(element: Element): Paragraph | undefined {
	if (element.kind === 'paragraph') return element;
	if (element.kind !== 'outline' || element.children.length !== 1) return undefined;
	const child = element.children[0];
	return child.kind === 'paragraph' ? child : undefined;
}

function highlightedConfigBlock(elements: Element[], start: number): { lines: string[], end: number } | undefined {
	const lines: string[] = [];
	let highlighted = 0;
	let syntax = 0;
	let end = start;
	for (; end < elements.length; end++) {
		const paragraph = simpleParagraph(elements[end]);
		if (!paragraph || paragraph.list || paragraph.styleId || paragraph.children.length > 0) break;
		const text = rawParagraphText(paragraph);
		if (text.trim() === '') {
			lines.push('');
			continue;
		}
		if (!redHighlightedParagraph(paragraph)) break;
		highlighted++;
		if (/^\s*(?:#|\[[^\]]+\]|[A-Za-z_][\w.-]*\s*=|\d+\.\s+\S)/.test(text)) syntax++;
		lines.push(text);
	}
	while (lines.at(-1) === '') lines.pop();
	return highlighted >= 8 && syntax / highlighted >= 0.7 ? { lines, end } : undefined;
}

function visitParagraphs(elements: Element[], visit: (paragraph: Paragraph) => void): void {
	for (const element of elements) {
		if (element.kind === 'paragraph') {
			visit(element);
			visitParagraphs(element.children, visit);
		}
		else if (element.kind === 'outline') visitParagraphs(element.children, visit);
	}
}

function codeLines(elements: Element[]): string[] {
	const lines: string[] = [];
	visitParagraphs(elements, paragraph => lines.push(...rawParagraphText(paragraph).split('\n')));
	return lines;
}

function monospaceRatio(elements: Element[]): number {
	let all = 0;
	let monospace = 0;
	visitParagraphs(elements, paragraph => {
		for (const run of paragraph.runs) {
			all += run.text.length;
			if (MONOSPACE_FONTS.has(run.font?.toLocaleLowerCase() ?? '')) monospace += run.text.length;
		}
	});
	return all === 0 ? 0 : monospace / all;
}

function tableColumns(table: Table): number {
	return Math.max(0, ...table.rows.map(row => row.cells.length));
}

function isLineNumberColumn(lines: string[]): boolean {
	const numbers = lines.map(line => Number(line.trim()));
	return numbers.length >= 2
		&& numbers.every(Number.isInteger)
		&& numbers.every((number, index) => index === 0 || number === numbers[index - 1] + 1);
}

function isEmptyColumn(lines: string[]): boolean {
	return lines.length > 0 && lines.every(line => line.trim() === '');
}

function withoutTrailingBlankLines(lines: string[]): string[] {
	let end = lines.length;
	while (end > 0 && lines[end - 1].trim() === '') end--;
	return lines.slice(0, end);
}

function codeFromTable(table: Table): string[] | undefined {
	const columns = tableColumns(table);
	if (columns === 1) {
		const elements = table.rows.flatMap(row => row.cells[0]?.children ?? []);
		const lines = codeLines(elements);
		return lines.length >= 2 && (monospaceRatio(elements) >= 0.8 || codeLanguage(lines) !== '') ? lines : undefined;
	}
	if (columns !== 2 || table.rows.length !== 1) return undefined;

	const [numbers, code] = table.rows[0].cells;
	const numberLines = codeLines(numbers?.children ?? []);
	const rawCodeText = codeLines(code?.children ?? []);
	const trimmedCodeText = withoutTrailingBlankLines(rawCodeText);
	if (isEmptyColumn(numberLines)) return trimmedCodeText.length >= 2 ? trimmedCodeText : undefined;
	// A one-row number/code table is a common OneNote command container.
	const visibleNumbers = numberLines.filter(line => line.trim() !== '');
	if (visibleNumbers.length === 1 && /^\d+\s*$/.test(visibleNumbers[0]) && codeLanguage(trimmedCodeText) === 'bash') return trimmedCodeText;
	if (!isLineNumberColumn(numberLines)) return undefined;
	if (rawCodeText.length === numberLines.length) return rawCodeText;
	if (trimmedCodeText.length === numberLines.length) return trimmedCodeText;
	return Math.abs(trimmedCodeText.length - numberLines.length) <= 2 && codeLanguage(trimmedCodeText) !== ''
		? trimmedCodeText
		: undefined;
}

function codeLanguage(lines: string[]): string {
	const text = lines.join('\n').replace(/\*\*/g, '');
	// VCF metadata starts with Markdown heading markers. Treating it as prose
	// turns an otherwise intact genomics file into hundreds of headings.
	if (/^##fileformat=VCFv[\d.]+\s*$/mi.test(text) || /^#CHROM\s+POS\s+ID\s+REF\s+ALT\b/m.test(text)) return 'vcf';
	if (!/^\s*[A-Za-z_]\w*\s*=.*?,\s*(?:which means|that is|i\.e\.)\b/mi.test(text)
		&& /^\s*(?:from\s+\S+\s+import|import\s+(?:pandas|numpy|os|sys|glob|gzip|h5py|Bio)\b|def\s+\w+\s*\(|class\s+\w+\s*[:(])/m.test(text)) return 'python';
	if (/^\s*(?:(?:library|require)\s*\(|.*<-|.*\bggplot\s*\(|.*\bread\.(?:csv|table)\s*\()/m.test(text)) return 'r';
	if (/^\s*(?:#!.*\b(?:ba)?sh\b|SBATCH\b|(?:if|for|while)\s+.+;\s*(?:then|do)\b|(?:cd|sed|awk|grep|find|scp|rsync|chmod|sbatch|(?:\S+\/)?(?:mapper|quantifier)\.pl)\s)/m.test(text)) return 'bash';
	if (/^\s*(?:param\s*\(|Get-|Set-|New-|Remove-|Write-Host\b|\$[A-Za-z_]\w*\s*=)/mi.test(text)) return 'powershell';
	if (/^\s*(?:select|insert\s+into|update\s+\S+\s+set|create\s+table|with\s+\w+\s+as\s*\()/mi.test(text)) return 'sql';
	if (/^(?:\s*\{[\s\S]*\}|\s*\[[\s\S]*])\s*$/.test(text) && /"[^"\n]+"\s*:/.test(text)) return 'json';
	if (/^(?:---\s*$|\s*[\w.-]+:\s+\S+)/m.test(text) && !/[;{}]/.test(text)) return 'yaml';
	return '';
}

function fencedCode(lines: string[]): string {
	const code = lines.join('\n').replace(/\s+$/g, '');
	const longest = Math.max(0, ...[...code.matchAll(/`+/g)].map(match => match[0].length));
	const fence = '`'.repeat(Math.max(3, longest + 1));
	return `${fence}${codeLanguage(lines)}\n${code}\n${fence}`;
}

function cleanHighlightedConfig(markdown: string): string {
	const lines = markdown.split('\n');
	const output: string[] = [];
	for (let index = 0; index < lines.length;) {
		if (!/^==🔴.*==$/.test(lines[index])) {
			output.push(lines[index++]);
			continue;
		}
		const collected: string[] = [];
		let highlighted = 0;
		let syntax = 0;
		let end = index;
		for (; end < lines.length; end++) {
			const line = lines[end];
			if (line === '') {
				collected.push('');
				continue;
			}
			const match = /^==🔴([\s\S]*)==$/.exec(line);
			if (!match) break;
			const text = match[1];
			highlighted++;
			if (/^\s*(?:#|\[[^\]]+\]|[A-Za-z_][\w.-]*\s*=|\d+\.\s+\S)/.test(text)) syntax++;
			collected.push(text);
		}
		while (collected.at(-1) === '') collected.pop();
		if (highlighted >= 8 && syntax / highlighted >= 0.7) {
			output.push(fencedCode(collected));
			index = end;
		}
		else output.push(lines[index++]);
	}
	return output.join('\n');
}

/**
 * Some OneNote versions wrap each VCF line in a separate outline, bypassing
 * the semantic paragraph collector. At this point the text is already exact;
 * only fence it so Markdown cannot reinterpret VCF metadata as headings.
 */
function fenceVcfMarkdown(markdown: string): string {
	const lines = markdown.split('\n');
	const start = lines.findIndex(line => /^##fileformat=VCFv[\d.]+\s*$/i.test(line));
	if (start === -1 || lines.slice(start).some(line => /^```/.test(line))) return markdown;
	return [...lines.slice(0, start), '```vcf', ...lines.slice(start), '```'].join('\n');
}

/** A long Python dict pasted as text must not be interpreted as prose. */
function fencePythonDictMarkdown(markdown: string): string {
	const lines = markdown.split('\n');
	const output: string[] = [];
	for (const line of lines) {
		if (/^\{['"][A-Za-z_]\w*['"]:\s*[\s\S]*\}\s*$/.test(line) && line.length >= 200) {
			output.push('```python', cleanCodeLikeLine(line, 'python'), '```');
		}
		else output.push(line);
	}
	return output.join('\n');
}

/**
 * Some desktop OneNote revisions expose an Office HYPERLINK field as literal
 * text instead of a TextRun hyperlink. Keep its visible label and turn normal
 * destinations into Markdown links; OneNote destinations deliberately become
 * plain text when their target cannot be resolved at this stage.
 */
function normalizeOfficeHyperlinkFields(markdown: string): string {
	return markdown.replace(/[\p{Co}\uFDEF\s]*HYPERLINK "([^"]+)"([^\r\n|]*)/giu, (whole, destination: string, visible: string) => {
		const label = visible.trim();
		if (destination.toLowerCase().startsWith('onenote:')) return label;
		return label === '' ? whole : `[${label}](${encodeURI(destination)})`;
	});
}

function cleanXmlLine(line: string): string {
	return line
		.replace(/\*\*/g, '')
		.replace(/^\*(<!--[\s\S]*-->)\*$/u, '$1')
		.replace(/\\+([<>?\x5b\x5d])/g, '$1');
}

/** Fence literal XML that OneNote split into richly formatted text runs. */
function fenceXmlMarkdown(markdown: string): string {
	const lines = markdown.split('\n');
	const output: string[] = [];
	for (let index = 0; index < lines.length;) {
		if (!/^\s*<\?xml(?:\s|\?|$)/i.test(cleanXmlLine(lines[index]))) {
			output.push(lines[index++]);
			continue;
		}
		const xml: string[] = [];
		let end = index;
		for (; end < lines.length; end++) {
			const cleaned = cleanXmlLine(lines[end]);
			if (cleaned.trim() !== '' && !/^\s*</u.test(cleaned)) break;
			xml.push(cleaned);
		}
		if (xml.filter(line => line.trim() !== '').length >= 3) {
			output.push('```', ...xml, '```');
			index = end;
		}
		else output.push(lines[index++]);
	}
	return output.join('\n');
}

/**
 * A few OneNote code containers survive semantic parsing as a 1x2 GFM table:
 * a single line number and one shell command. Keep real tables untouched.
 */
function fenceNumberedShellTable(markdown: string): string {
	const lines = markdown.split('\n');
	const output: string[] = [];
	for (let index = 0; index < lines.length;) {
		const row = /^\\?\|\s*\d+\s*\\?\|\s*(.*?)\s*\\?\|\s*$/u.exec(lines[index]);
		const separator = /^\\?\|\s*:?-+:?\s*\\?\|\s*:?-+:?\s*\\?\|\s*$/u.test(lines[index + 1] ?? '');
		if (row && separator && codeLanguage([row[1]]) === 'bash') {
			output.push('```bash', row[1], '```');
			index += 2;
		}
		else output.push(lines[index++]);
	}
	return output.join('\n');
}

type FencedLanguage = 'bash' | 'python' | 'r';

interface CodeLikeLine {
	language: FencedLanguage | 'comment';
	text: string;
}

/** Removes Markdown added while rendering syntax-coloured code, never prose. */
function cleanCodeLikeLine(line: string, language: FencedLanguage | 'comment'): string {
	let cleaned = line
		.replace(/==[🔴🟠🟡🟢🔵🟣]([\s\S]*?)==/gu, '$1')
		.replace(/^\s*\\#/, '#')
		.replace(/^\*#([\s\S]*?)\*$/, '#$1')
		.replace(/\*#([^*\r\n]+)\*/g, ' #$1')
		.replace(/\\(?:\[|\])/g, match => match.slice(1))
		.replace(/\s{2}$/, '');
	if (language === 'bash') return cleaned.replace(/\*\*/g, '');
	return cleaned.replace(/\*\*([A-Za-z_][\w-]*)\*\*/g, '$1');
}

function gffLine(line: string): string | undefined {
	const cleaned = cleanCodeLikeLine(line, 'comment').trim();
	return /^\S+\s+\S+\s+(?:gene|mRNA|transcript|exon|CDS|five_prime_UTR|three_prime_UTR)\s+\d+\s+\d+\s+\S+\s+[.+-]\s+\S+\s+\S+=/u.test(cleaned)
		? cleaned
		: undefined;
}

function codeLikeLine(line: string): CodeLikeLine | undefined {
	const cleaned = cleanCodeLikeLine(line.replace(/\*\*/g, ''), 'comment').trim();
	if (cleaned === '') return undefined;
	if (/^#(?:!|\$|\s|[A-Za-z])/u.test(cleaned)) return { language: 'comment', text: cleaned };
	if (/^[A-Za-z_]\w*\s*=.*?,\s*(?:which means|that is|i\.e\.)\b/iu.test(cleaned)) return undefined;
	if (/^(?:(?:library|require)\s*\(|[A-Za-z_.][\w.]*\s*<-|[A-Za-z_.][\w.]*\s*=\s*(?:find\.package|paste|tempfile|numeric|SlicedData\$new)\b|ggplot\s*\(|(?:read|write)\.(?:csv|table)\s*\()/u.test(cleaned)) {
		return { language: 'r', text: cleanCodeLikeLine(line, 'r') };
	}
	if (/^(?:from\s+\S+\s+import\b|import\s+\S+|def\s+\w+\s*\(|class\s+\w+|(?:elif|else|try|except|finally)\b|for\s+.+\s+in\s+.+:|while\s+.+:|if\s+.+:|with\s+.+:|return\b|pass\b|break\b|continue\b|[A-Za-z_]\w*(?:\[[^\]]+\]|\.\w+)*\s*=)/u.test(cleaned)) {
		return { language: 'python', text: cleanCodeLikeLine(line, 'python') };
	}
	if (/^(?:#!|#\$|(?:cd|wget|curl|gzip|gunzip|java|gatk|samtools|bwa|bcftools|bedtools|(?:\S+\/)?(?:mapper|quantifier)\.pl|python(?:\d+(?:\.\d+)?)?|perl|Rscript|rsync|scp|ssh|mkdir|rm|cp|mv|find|grep|awk|sed|sort|uniq|sbatch|qsub|module)\b|(?:export\s+)?[A-Za-z_]\w*=|(?:if|for|while)\s+.+;\s*(?:then|do)\b)/u.test(cleaned)) {
		return { language: 'bash', text: cleanCodeLikeLine(line, 'bash') };
	}
	return undefined;
}

function standaloneBashCommand(line: string): boolean {
	const cleaned = cleanCodeLikeLine(line, 'bash').trim();
	return /^(?:for\s+.+;\s*do\b[\s\S]*;\s*done\s*$|(?:cd|wget|curl|gzip|gunzip|java|gatk|samtools|bwa|bcftools|bedtools|(?:\S+\/)?(?:mapper|quantifier)\.pl|python(?:\d+(?:\.\d+)?)?|perl|Rscript|rsync|scp|ssh|mkdir|rm|cp|mv|find|grep|awk|sed|sort|uniq|sbatch|qsub|module)\s+\S+)/u.test(cleaned)
		&& (cleaned.length >= 60 || /(?:\s--?[\w-]+|\s\/|\s\|\s|;\s*(?:do|done)\b)/u.test(cleaned));
}

/** A less strict continuation is safe only after a code block has a proven start. */
function codeLikeContinuation(line: string, language: FencedLanguage): string | undefined {
	const cleaned = cleanCodeLikeLine(line, language);
	const trimmed = cleaned.trim();
	if (trimmed === '') return '';
	if (trimmed.startsWith('#')) return trimmed;
	if (codeLikeLine(line)?.language === language) return cleaned;
	if (language === 'python' && /^(?:\s+|[A-Za-z_]\w*(?:\[[^\]]+\]|\.\w+)*(?:\s*(?:\+=|-=|\*=|\/=|=)|\([^\r\n]*\))|(?:\[|\]|\{|\}|\(|\)|,|\.|'|"|`)|else:|try:|except\b)/u.test(cleaned)) return cleaned;
	if (language === 'r' && /^(?:\s+|[A-Za-z_.]\w*(?:\$\w+)?\s*(?:<-|=|\()|(?:\[|\]|\{|\}|\(|\)|,|\.|'|"|`))/u.test(cleaned)) return cleaned;
	if (language === 'bash' && /^(?:\s+|(?:then|do|done|fi|else|elif)\b|[|>&])/u.test(cleaned)) return cleaned;
	return undefined;
}

function fenceStructuredMarkdown(markdown: string): string {
	const lines = markdown.split('\n');
	const output: string[] = [];
	for (let index = 0; index < lines.length;) {
		// Earlier semantic passes already made these code blocks. Never inspect
		// their contents as prose or add a second fence around them.
		if (/^```/.test(lines[index])) {
			const end = lines.findIndex((line, candidate) => candidate > index && /^```/.test(line));
			if (end < 0) return markdown;
			output.push(...lines.slice(index, end + 1));
			index = end + 1;
			continue;
		}
		const firstGff = gffLine(lines[index]);
		if (firstGff) {
			const records = [firstGff];
			let end = index + 1;
			for (; end < lines.length; end++) {
				const record = gffLine(lines[end]);
				if (record) records.push(record);
				else if (lines[end].trim() !== '') break;
			}
			if (records.length >= 3) {
				output.push('```gff', ...records, '```');
				index = end;
				continue;
			}
		}

		const first = codeLikeLine(lines[index]);
		if (!first || first.language === 'comment') {
			output.push(lines[index++]);
			continue;
		}
		const language = first.language;
		const code: string[] = [first.text];
		let strong = 1;
		let end = index + 1;
		for (; end < lines.length; end++) {
			if (lines[end].trim() === '') {
				continue;
			}
			const next = codeLikeLine(lines[end]);
			if (next?.language === language || next?.language === 'comment') {
				code.push(next.text);
				if (next.language === language) strong++;
				continue;
			}
			const continuation = codeLikeContinuation(lines[end], language);
			if (continuation === undefined) break;
			code.push(continuation);
		}
		const explicitLongCommand = language === 'bash' && code.length === 1 && standaloneBashCommand(code[0]);
		if (strong >= 2 || explicitLongCommand) {
			output.push(`\`\`\`${language}`, ...code, '```');
			index = end;
		}
		else output.push(lines[index++]);
	}
	return output.join('\n');
}

function html(value: string): string {
	const cleaned = [...value].map(character => {
		const code = character.charCodeAt(0);
		if (code === 0x0b) return '\n';
		return code <= 0x08 || code === 0x0c || (code >= 0x0e && code <= 0x1f) || code === 0x7f ? '' : character;
	}).join('');

	return cleaned
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/\r?\n/g, '<br>');
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
	if (Math.max(red, green, blue) - Math.min(red, green, blue) <= 12) return text;

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

function cleanHeadingText(text: string): string {
	return text
		.replace(/\s*\\\[\[\*?(?:Top of page|页面顶部)\*?\]\([^)]+\)\\\]/giu, '')
		.replace(/^(\*{1,3})?\\?#{1,6}\s+/, '$1');
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

function imageLabel(fileName: string | undefined): string | undefined {
	if (!fileName) return undefined;
	const baseName = fileName.split(/[\\/]/).at(-1)?.trim().replace(/[\r\n\v]+/g, ' ');
	if (!baseName) return undefined;
	return escapeInline(baseName.replace(/\\/g, '\\\\'));
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

interface RenderContext {
	listDepth: number;
	parentList?: {
		level: number;
		prefix: string;
		hasText: boolean;
	};
}

const ROOT_CONTEXT: RenderContext = { listDepth: 0 };

function extensionOf(fileName: string | undefined): string | undefined {
	return fileName?.match(/\.[^.\\/]+$/)?.[0];
}

/** An attachment without an extension is one the vault cannot open. */
function withExtension(base: string, extension: string | undefined): string {
	if (!extension) return base;
	const normalized = extension.startsWith('.') ? extension : `.${extension}`;
	// A page title can itself end in a dot and ordinary words (for example,
	// "Dr. Matt Hurles" or a protein version ending in ".1").  That is not an
	// attachment extension.  Only leave a name alone when it already has the
	// exact extension supplied by OneNote.
	if (extensionOf(base)?.toLocaleLowerCase() === normalized.toLocaleLowerCase()) return base;
	return base + normalized;
}

class PageWriter {
	private readonly blocks: Block[] = [];
	private readonly inkStrokes: SvgStroke[] = [];
	private readonly imageOrdinals = new Map<string, number>();
	readonly attachments: ResolvedAttachment[] = [];
	readonly missingAssets: MissingAsset[] = [];
	htmlFallbacks = 0;
	cancelled = false;

	constructor(private readonly options: OneNoteConversionOptions, private readonly pageTitle: string) {}

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

	async writeElements(elements: Element[], context: RenderContext = ROOT_CONTEXT): Promise<void> {
		for (let index = 0; index < elements.length;) {
			if (this.options.isCancelled?.()) {
				this.cancelled = true;
				return;
			}
			const code = highlightedConfigBlock(elements, index);
			if (code) {
				this.push(fencedCode(code.lines));
				index = code.end;
				continue;
			}
			await this.writeElement(elements[index], context);
			index++;
		}
	}

	private async writeElement(element: Element, context: RenderContext): Promise<void> {
		switch (element.kind) {
			case 'outline': {
				const level = element.list ? Math.max(element.list.level, context.listDepth) : 0;
				const list = element.list && { ...element.list, level };
				await this.writeElements(element.children, {
					...context,
					listDepth: element.list ? level + 1 : context.listDepth,
					parentList: list ? {
						level,
						prefix: listPrefix(list),
						hasText: false,
					} : context.parentList,
				});
				break;
			}
			case 'paragraph':
				await this.writeParagraph(element, context);
				break;
			case 'table':
				await this.writeTable(element);
				break;
			case 'image':
				await this.writeImage(element, context);
				break;
			case 'embedded-file': {
				const name = withExtension(element.fileName ?? 'attachment', element.extension);
				await this.writeAsset(element.data, name, name, false, context);
				break;
			}
			case 'ink':
				this.collectInk(element);
				break;
		}
	}

	private async writeParagraph(paragraph: Paragraph, context: RenderContext): Promise<void> {
		const raw = rawParagraphText(paragraph);
		if (!paragraph.list && !paragraph.styleId
			&& raw.length >= 120
			&& monospaceRatio([paragraph]) >= 0.8
			&& !/^\s*[A-Za-z_]\w*\s*=.*?,\s*(?:which means|that is|i\.e\.)\b/mi.test(raw)
			&& /^(?:\s*[A-Za-z_]\w*(?:\[[^\n]+\])?\s*=|\s*(?:for|if|while)\b|\s*\w+(?:\.\w+)+\s*\()/m.test(raw)) {
			this.push(fencedCode(raw.split('\n')));
			await this.writeElements(paragraph.children, context);
			return;
		}
		let text = renderRuns(paragraph.runs, this.options);
		const level = paragraph.list ? Math.max(paragraph.list.level, context.listDepth) : 0;
		const list = paragraph.list && { ...paragraph.list, level };
		const task = taskPrefix(paragraph.tags, list);
		const prefix = task ?? listPrefix(list);
		let wroteText = false;

		if (text !== '') {
			const indent = '\t'.repeat(level);
			const heading = headingPrefix(paragraph.styleId);
			if (heading) text = cleanHeadingText(text);

			const escaped = text.split('\n').map(escapeLineStart);
			const body = (prefix || heading) + escaped.join('  \n' + indent);
			const callout = calloutFor(paragraph.tags);

			if (callout && !paragraph.list && !task) this.pushCallout(callout, body);
			else this.push(body, task !== undefined || paragraph.list !== undefined);
			wroteText = true;
		}

		await this.writeElements(paragraph.children, paragraph.list ? {
			listDepth: level + 1,
			parentList: { level, prefix, hasText: wroteText },
		} : context);
	}

	private async writeTable(table: Table): Promise<void> {
		if (table.rows.length === 0) return;
		const code = codeFromTable(table);
		if (code) {
			this.push(fencedCode(code));
			return;
		}

		const columns = tableColumns(table);
		if (columns === 1) {
			for (const row of table.rows) await this.writeElements(row.cells[0]?.children ?? []);
			return;
		}

		if (this.containsComplexNestedTable(table)) {
			this.htmlFallbacks++;
			this.push(await this.renderHtmlTable(table));
			return;
		}

		const rendered: string[][] = [];
		for (const row of table.rows) {
			const cells: string[] = [];
			for (let index = 0; index < columns; index++) {
				const text = await this.renderCell(row.cells[index]?.children ?? [], '<br>');
				cells.push(text.replace(/\|/g, '\\|').trim());
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

	}

	async writeCollectedInk(): Promise<void> {
		const svg = strokesToSvg(this.inkStrokes);
		if (!svg) return;

		await this.writeAsset(new TextEncoder().encode(svg), `${this.pageTitle} - Ink.svg`, undefined, true, ROOT_CONTEXT);
	}

	private imageName(image: Image): string {
		return withExtension(`${this.pageTitle} image`, image.extension ?? extensionOf(image.fileName));
	}

	private imageSource(image: Image): AssetSource {
		const sourceName = image.fileName;
		const key = sourceName ?? '';
		const ordinal = this.imageOrdinals.get(key) ?? 0;
		this.imageOrdinals.set(key, ordinal + 1);
		return { sourceName, ordinal, embed: true };
	}

	private async writeImage(image: Image, context: RenderContext): Promise<void> {
		await this.writeAsset(image.data, this.imageName(image), imageLabel(image.fileName), true, context, this.imageSource(image));
	}

	private async writeAsset(data: Uint8Array | undefined, name: string, label: string | undefined, embed: boolean, context: RenderContext, source?: AssetSource): Promise<void> {
		const link = await this.renderAsset(data, name, label, embed, source);
		if (!link) return;
		const parentList = context.parentList;
		if (!parentList) this.push(link);
		else if (parentList.hasText) this.push(`${'\t'.repeat(parentList.level + 1)}${link}`, true);
		else this.push(parentList.prefix + link, true);
	}

	private async renderAsset(data: Uint8Array | undefined, name: string, label: string | undefined, embed: boolean, source?: AssetSource): Promise<string | undefined> {
		if (!data || data.length === 0) {
			this.options.onSkipped?.(name, 'no-data');
			this.missingAssets.push({ name, label: label ?? name, embed, ...source });
			return undefined;
		}

		const attachment = await this.options.saveAttachment(data, name, source);
		if (!attachment) {
			this.options.onSkipped?.(name, 'no-data');
			this.missingAssets.push({ name, label: label ?? name, embed, ...source });
			return undefined;
		}

		this.attachments.push({ ...attachment, ...source });
		const target = encodeURI(attachment.path);
		const display = label ?? imageLabel(attachment.name) ?? '';
		return embed ? `![${display}](${target})` : `[${display}](${target})`;
	}

	private async renderCell(children: Element[], separator = ' '): Promise<string> {
		const parts: string[] = [];

		for (const child of children) {
			switch (child.kind) {
				case 'paragraph':
					parts.push(renderRuns(child.runs, this.options));
					parts.push(await this.renderCell(child.children, separator));
					break;
				case 'outline':
					parts.push(await this.renderCell(child.children, separator));
					break;
				case 'image':
					parts.push(await this.renderAsset(child.data, this.imageName(child), imageLabel(child.fileName), true, this.imageSource(child)) ?? '');
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
					parts.push(await this.renderNestedTable(child));
					break;
			}
		}

		return parts.filter(part => part !== '').join(separator);
	}

	private containsComplexNestedTable(table: Table): boolean {
		const contains = (elements: Element[]): boolean => elements.some(element => {
			if (element.kind === 'table') return tableColumns(element) > 1 || this.containsComplexNestedTable(element);
			return (element.kind === 'paragraph' || element.kind === 'outline') && contains(element.children);
		});
		return table.rows.some(row => row.cells.some(cell => contains(cell.children)));
	}

	private async renderNestedTable(table: Table): Promise<string> {
		const code = codeFromTable(table);
		if (code) return fencedCode(code);
		if (tableColumns(table) === 1) {
			const items: string[] = [];
			for (const row of table.rows) {
				const value = await this.renderCell(row.cells[0]?.children ?? [], '\n');
				if (value.trim() !== '') items.push(value.split('\n').map((line, index) => `${index === 0 ? '- ' : '  '}${line}`).join('\n'));
			}
			return items.join('\n');
		}
		this.htmlFallbacks++;
		return await this.renderHtmlTable(table);
	}

	private async renderHtmlElements(elements: Element[]): Promise<string> {
		const parts: string[] = [];
		for (const element of elements) {
			if (element.kind === 'paragraph') {
				const text = element.runs.map(run => run.text).join('');
				if (text !== '') parts.push(`<p>${html(text)}</p>`);
				parts.push(await this.renderHtmlElements(element.children));
			}
			else if (element.kind === 'outline') parts.push(await this.renderHtmlElements(element.children));
			else if (element.kind === 'table') parts.push(await this.renderHtmlTable(element));
			else if (element.kind === 'image') {
				const link = await this.renderAsset(element.data, this.imageName(element), imageLabel(element.fileName), true, this.imageSource(element));
				if (link) {
					const match = /^!\[(.*)\]\((.*)\)$/.exec(link);
					if (match) parts.push(`<img src="${html(match[2])}" alt="${html(match[1])}">`);
				}
			}
			else if (element.kind === 'embedded-file') {
				const name = withExtension(element.fileName ?? 'attachment', element.extension);
				const link = await this.renderAsset(element.data, name, name, false);
				const match = link && /^\[(.*)\]\((.*)\)$/.exec(link);
				if (match) parts.push(`<a href="${html(match[2])}">${html(match[1])}</a>`);
			}
			else this.collectInk(element);
		}
		return parts.filter(Boolean).join('');
	}

	private async renderHtmlTable(table: Table): Promise<string> {
		const rows: string[] = [];
		for (const row of table.rows) {
			const cells: string[] = [];
			for (const cell of row.cells) cells.push(`<td>${await this.renderHtmlElements(cell.children)}</td>`);
			rows.push(`<tr>${cells.join('')}</tr>`);
		}
		return `<table>\n${rows.join('\n')}\n</table>`;
	}
}

export async function convertPage(page: Page, options: OneNoteConversionOptions): Promise<ConvertedPage> {
	const writer = new PageWriter(options, options.noteName ?? page.title);

	await writer.writeElements(page.outlines);
	await writer.writeElements(page.directContent);
	await writer.writeCollectedInk();

	return {
		markdown: fenceStructuredMarkdown(fencePythonDictMarkdown(fenceVcfMarkdown(fenceNumberedShellTable(fenceXmlMarkdown(normalizeOfficeHyperlinkFields(cleanHighlightedConfig(writer.markdown))))))),
		attachments: writer.attachments,
		missingAssets: writer.missingAssets,
		htmlFallbacks: writer.htmlFallbacks,
		degraded: writer.missingAssets.length > 0,
		cancelled: writer.cancelled,
	};
}
