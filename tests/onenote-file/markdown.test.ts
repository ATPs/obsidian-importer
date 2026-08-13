import { test } from 'node:test';
import assert from 'node:assert/strict';

import { convertPage } from '../../src/formats/onenote-file/convert';
import type { Element, Page, Paragraph, TextRun } from '../../src/formats/onenote-file/semantic/content';

function page(...elements: Element[]): Page {
	return {
		id: 'test-page',
		title: 'Test',
		level: 0,
		isConflictPage: false,
		isDeleted: false,
		outlines: [{ kind: 'outline', children: elements }],
		directContent: [],
		preservation: [],
	};
}

function para(runs: TextRun[] | string, extra: Partial<Paragraph> = {}): Paragraph {
	return {
		kind: 'paragraph',
		runs: typeof runs === 'string' ? [{ text: runs }] : runs,
		children: [],
		...extra,
	};
}

async function render(...elements: Element[]): Promise<string> {
	const converted = await convertPage(page(...elements), {
		saveAttachment: async (_data, name) => ({ path: `files/${name}`, name }),
	});
	return converted.markdown;
}

test('emphasis wraps the words and not the spaces around them', async () => {
	assert.equal(await render(para([{ text: 'a ' }, { text: 'bold ', bold: true }, { text: 'b' }])), 'a **bold** b');
	assert.equal(await render(para([{ text: ' padded ', italic: true }])), '*padded*');
});

test('a run can carry more than one kind of emphasis', async () => {
	assert.equal(await render(para([{ text: 'x', bold: true, italic: true }])), '***x***');
	assert.equal(await render(para([{ text: 'y', bold: true, strikethrough: true }])), '~~**y**~~');
});

test('a link wraps whatever emphasis the run already had', async () => {
	assert.equal(
		await render(para([{ text: 'site', bold: true, hyperlinkUrl: 'https://example.com/a b' }])),
		'[**site**](https://example.com/a%20b)');
});

test('a OneNote page link becomes a link Obsidian can resolve after import', async () => {
	const converted = await convertPage(page(para([{
		text: 'the other page',
		hyperlinkUrl: 'onenote:///C:/Notebook/Section.one#Other%20page&section-id={section}&page-id={page}&end',
	}])), { saveAttachment: async () => null, resolveInternalLink: title => title });
	assert.equal(converted.markdown, '[the other page](Other%20page)');
});

test('a double-encoded OneNote page title still resolves to its imported page', async () => {
	const converted = await convertPage(page(para([{
		text: 'the other page',
		hyperlinkUrl: 'onenote:///C:/Notebook/Section.one#Other%2520page&section-id={section}&page-id={page}&end',
	}])), { saveAttachment: async () => null, resolveInternalLink: title => title });
	assert.equal(converted.markdown, '[the other page](Other%20page)');
});

test('a malformed OneNote page title is kept rather than failing its page', async () => {
	const converted = await convertPage(
		page(para([{ text: 'page', hyperlinkUrl: 'onenote:///Section.one#Bad%escape&section-id={section}' }])),
		{ saveAttachment: async () => null, resolveInternalLink: title => title });
	assert.equal(converted.markdown, '[page](Bad%25escape)');
});

test('an unresolved OneNote page link keeps only its visible text', async () => {
	const url = 'onenote:///S.one#Duplicate&section-id={a}&page-id={b}';
	const converted = await convertPage(page(para([{ text: 'page', hyperlinkUrl: url }])), {
		saveAttachment: async () => null,
		resolveInternalLink: () => undefined,
	});
	assert.equal(converted.markdown, 'page');
});

test('an ambiguous OneNote page link keeps only its visible text', async () => {
	const converted = await convertPage(page(para([{
		text: 'duplicate page',
		hyperlinkUrl: 'onenote:///S.one#Duplicate&section-id={a}',
	}])), {
		saveAttachment: async () => null,
		resolveInternalLink: () => ['Section A/Duplicate abc123', 'Section B/Duplicate def456'],
	});

	assert.equal(converted.markdown, 'duplicate page');
});

test('a run of only whitespace contributes no emphasis markers', async () => {
	assert.equal(await render(para([{ text: '   ', bold: true }])), '');
});

test('a style identifier becomes a heading', async () => {
	assert.equal(await render(para('Title', { styleId: 'h1' })), '# Title');
	assert.equal(await render(para('Deep', { styleId: 'h6' })), '###### Deep');
	assert.equal(await render(para('Body', { styleId: 'p' })), 'Body');
});

test('a heading drops duplicate markers and OneMore top-of-page navigation', async () => {
	const url = 'onenote:///S.one#Test&section-id={a}';
	const converted = await convertPage(page(para([
		{ text: '## ' },
		{ text: 'results', bold: true },
		{ text: ' [' },
		{ text: 'Top of page', italic: true, hyperlinkUrl: url },
		{ text: ']' },
	], { styleId: 'h2' })), {
		saveAttachment: async () => null,
		resolveInternalLink: () => 'Test',
	});

	assert.equal(converted.markdown, '## **results**');
});

test('lists carry their bullet and their indent', async () => {
	const markdown = await render(
		para('one', { list: { level: 0, ordered: false } }),
		para('nested', { list: { level: 1, ordered: false } }),
		para('numbered', { list: { level: 0, ordered: true } }));

	assert.equal(markdown, ['- one', '\t- nested', '1. numbered'].join('\n'));
});

test('a list is separated from the prose around it', async () => {
	const markdown = await render(
		para('before'),
		para('one', { list: { level: 0, ordered: false } }),
		para('two', { list: { level: 0, ordered: false } }),
		para('after'));

	assert.equal(markdown, ['before', '', '- one', '- two', '', 'after'].join('\n'));
});

test('an item that is not a list item breaks the list', async () => {
	const markdown = await render(
		para('one', { list: { level: 0, ordered: false } }),
		para('interrupting'),
		para('two', { list: { level: 0, ordered: false } }));

	assert.equal(markdown, ['- one', '', 'interrupting', '', '- two'].join('\n'));
});

test('a list item wins over a heading style on the same paragraph', async () => {
	assert.equal(await render(para('item', { list: { level: 0, ordered: false }, styleId: 'h2' })), '- item');
});

test('a table gets a header row because GFM has no table without one', async () => {
	const markdown = await render({
		kind: 'table',
		rows: [
			{ cells: [{ children: [para('A')] }, { children: [para('B')] }] },
			{ cells: [{ children: [para('1')] }, { children: [para('2')] }] },
		],
	});

	assert.equal(markdown, ['| A | B |', '| --- | --- |', '| 1 | 2 |'].join('\n'));
});

test('a ragged table is padded to its widest row', async () => {
	const markdown = await render({
		kind: 'table',
		rows: [
			{ cells: [{ children: [para('A')] }] },
			{ cells: [{ children: [para('1')] }, { children: [para('2')] }, { children: [para('3')] }] },
		],
	});

	assert.equal(markdown, ['| A |  |  |', '| --- | --- | --- |', '| 1 | 2 | 3 |'].join('\n'));
});

test('a non-code single-column table is expanded as ordinary content', async () => {
	const markdown = await render({
		kind: 'table',
		rows: [{ cells: [{ children: [para('a | b')] }] }],
	});

	assert.equal(markdown, 'a | b');
});

test('a table with no rows produces nothing at all', async () => {
	assert.equal(await render({ kind: 'table', rows: [] }), '');
});

test('an image is named after its page but labelled from its source filename', async () => {
	const data = new Uint8Array([1, 2, 3]);

	assert.equal(
		await render({ kind: 'image', fileName: 'Untitled picture.png', altText: 'a shot', data }),
		'![Untitled picture.png](files/Test%20image.png)');
	assert.equal(
		await render({ kind: 'image', extension: '.jpg', data }),
		'![Test image.jpg](files/Test%20image.jpg)');

	assert.equal(
		await render({ kind: 'embedded-file', fileName: 'notes.docx', data }),
		'[notes.docx](files/notes.docx)');
});

test('image labels do not expose OneNote recognition text', async () => {
	const converted = await convertPage(page({
		kind: 'image',
		fileName: 'diagram [final].png',
		altText: 'A very long OCR transcription\nthat must not appear in Markdown.',
		data: new Uint8Array([1]),
	}), { saveAttachment: async (_bytes, name) => ({ path: `files/${name}`, name }) });

	assert.equal(converted.markdown, '![diagram \\[final\\].png](files/Test%20image.png)');
	assert.doesNotMatch(converted.markdown, /OCR transcription/);
	assert.deepEqual(converted.attachments.map(asset => ({ sourceName: asset.sourceName, ordinal: asset.ordinal, embed: asset.embed })), [{
		sourceName: 'diagram [final].png', ordinal: 0, embed: true,
	}]);
});

test('images with identical source names have stable per-name ordinals', async () => {
	const converted = await convertPage(page(
		{ kind: 'image', fileName: 'plot.png', data: new Uint8Array([1]) },
		{ kind: 'image', fileName: 'plot.png', data: new Uint8Array([2]) },
	), { saveAttachment: async (_bytes, name, source) => ({ path: `files/${name}`, name, ...source }) });

	assert.deepEqual(converted.attachments.map(asset => asset.ordinal), [0, 1]);
});

test('an asset with no bytes is reported rather than linked', async () => {
	const skipped: string[] = [];

	const converted = await convertPage(page({ kind: 'image', fileName: 'gone.png' }), {
		saveAttachment: async () => assert.fail('an asset with no bytes should never be saved'),
		onSkipped: name => skipped.push(name),
	});

	assert.equal(converted.markdown, '');
	assert.deepEqual(skipped, ['Test image.png']);
	assert.deepEqual(converted.attachments, []);
});

test('an asset the importer refuses is left out of the markdown', async () => {
	const converted = await convertPage(page({ kind: 'image', fileName: 'big.png', data: new Uint8Array([1]) }), {
		saveAttachment: async () => null,
	});

	assert.equal(converted.markdown, '');
	assert.deepEqual(converted.attachments, []);
});

test('every saved attachment is reported back to the caller', async () => {
	const data = new Uint8Array([1]);

	const converted = await convertPage(
		page({ kind: 'image', fileName: 'a.png', data }, { kind: 'image', fileName: 'b.png', data }),
		{ saveAttachment: async (_bytes, name) => ({ path: `files/${name}`, name }) });

	assert.deepEqual(converted.attachments.map(attachment => attachment.name), ['Test image.png', 'Test image.png']);
});

test('a line break inside a paragraph stays inside it', async () => {
	assert.equal(await render(para('first\nsecond')), 'first  \nsecond');
});

test('empty paragraphs do not pile up blank lines', async () => {
	assert.equal(await render(para(''), para('text'), para('   '), para('more')), 'text\n\nmore');
});

test('children of a paragraph follow it', async () => {
	const markdown = await render(para('parent', { children: [para('child', { list: { level: 1, ordered: false } })] }));

	assert.equal(markdown, 'parent\n\n\t- child');
});

test('structural list nesting supplements missing local list levels', async () => {
	const markdown = await render(para('parent', {
		list: { level: 0, ordered: false },
		children: [para('child', {
			list: { level: 0, ordered: false },
			children: [para('grandchild', { list: { level: 0, ordered: true } })],
		})],
	}));

	assert.equal(markdown, ['- parent', '\t- child', '\t\t1. grandchild'].join('\n'));
});

test('an image owned by a list item stays in that list', async () => {
	const markdown = await render(para('figure', {
		list: { level: 0, ordered: false },
		children: [{ kind: 'image', fileName: 'chart.png', data: new Uint8Array([1]) }],
	}));

	assert.equal(markdown, ['- figure', '\t![chart.png](files/Test%20image.png)'].join('\n'));
});

test('a list item containing only an image uses the list marker', async () => {
	const markdown = await render(para('', {
		list: { level: 0, ordered: false },
		children: [{ kind: 'image', fileName: 'chart.png', data: new Uint8Array([1]) }],
	}));

	assert.equal(markdown, '- ![chart.png](files/Test%20image.png)');
});

test('a list wrapper around an image uses the wrapper marker and indent', async () => {
	const markdown = await render({
		kind: 'outline',
		list: { level: 0, ordered: false },
		children: [{ kind: 'image', fileName: 'chart.png', data: new Uint8Array([1]) }],
	});

	assert.equal(markdown, '- ![chart.png](files/Test%20image.png)');
});

test('cancelling stops the conversion where it stands', async () => {
	let checksBeforeCancelling = 2;

	const converted = await convertPage(page(para('first'), para('second')), {
		saveAttachment: async () => null,
		isCancelled: () => checksBeforeCancelling-- <= 0,
	});

	assert.equal(converted.markdown, 'first');
	assert.equal(converted.cancelled, true);
});

test('an equation becomes LaTeX rather than the glyphs OneNote stored', async () => {
	assert.equal(await render(para([{ text: '\u{1D44E}=\u{1D44F}', math: true }])), '$a=b$');
	assert.equal(await render(para([{ text: '\u{1D6FC}+\u{1D6FD}', math: true }])), '$α+β$');
	assert.equal(await render(para([{ text: '\u{1D400}\u{1D401}', math: true }])), '$AB$');
});

test('an equation keeps raised and lowered digits as LaTeX scripts', async () => {
	assert.equal(await render(para([{ text: 'x²', math: true }])), '$x^{2}$');
	assert.equal(await render(para([{ text: 'x₁₂', math: true }])), '$x_{12}$');
	assert.equal(await render(para([{ text: 'e⁻ⁿ', math: true }])), '$e^{-n}$');
});

test('the invisible operators a layout engine needs are dropped', async () => {
	assert.equal(await render(para([{ text: 'f⁡(x)', math: true }])), '$f(x)$');
	assert.equal(await render(para([{ text: '2⁢x', math: true }])), '$2x$');
});

test('emphasis is not wrapped around an equation', async () => {
	assert.equal(await render(para([{ text: '\u{1D44E}', math: true, italic: true, bold: true }])), '$a$');
});

test('an equation keeps the spacing around it', async () => {
	assert.equal(
		await render(para([{ text: 'where ' }, { text: '\u{1D465}', math: true }, { text: ' is odd' }])),
		'where $x$ is odd');
});

test('a math run holding nothing but spaces produces no delimiters', async () => {
	assert.equal(await render(para([{ text: '   ', math: true }])), '');
});

test('a OneNote to-do becomes a task, ticked or not', async () => {
	assert.equal(
		await render(para('buy milk', { tags: [{ checkable: true, completed: false }] })),
		'- [ ] buy milk');
	assert.equal(
		await render(para('done', { tags: [{ checkable: true, completed: true }] })),
		'- [x] done');
});

test('a to-do inside a list keeps its indent and loses the bullet', async () => {
	const markdown = await render(
		para('heading'),
		para('nested task', {
			list: { level: 2, ordered: false },
			tags: [{ checkable: true, completed: false }],
		}));

	assert.equal(markdown, 'heading\n\n\t\t- [ ] nested task');
});

test('consecutive tasks stay together like any other list', async () => {
	const markdown = await render(
		para('one', { tags: [{ checkable: true, completed: false }] }),
		para('two', { tags: [{ checkable: true, completed: true }] }));

	assert.equal(markdown, '- [ ] one\n- [x] two');
});

const tag = (shape: number, label?: string) => [{ checkable: false, completed: false, shape, label }];

test('a tag that means "pay attention" becomes the matching admonition', async () => {
	assert.equal(await render(para('look', { tags: tag(13, 'Important') })), '> [!important] Important\n> look');
	assert.equal(await render(para('why?', { tags: tag(15, 'Question') })), '> [!question] Question\n> why?');
	assert.equal(await render(para('run', { tags: tag(17, 'Critical') })), '> [!danger] Critical\n> run');
	assert.equal(await render(para('idea', { tags: tag(21, 'Idea') })), '> [!tip] Idea\n> idea');
});

test('a tag that merely categorises leaves its paragraph alone', async () => {
	assert.equal(await render(para('0800 1234', { tags: tag(109, 'Phone number') })), '0800 1234');
	assert.equal(await render(para('Dune', { tags: tag(132, 'Book to read') })), 'Dune');
	assert.equal(await render(para('a song', { tags: tag(121, 'Music to listen to') })), 'a song');
});

test('the title is the label as written, whatever language it is in', async () => {
	assert.equal(await render(para('merk dir das', { tags: tag(13, 'Wichtig') })), '> [!important] Wichtig\n> merk dir das');
});

test('paragraphs tagged the same way in a row become one admonition', async () => {
	const markdown = await render(
		para('first', { tags: tag(13, 'Important') }),
		para('second', { tags: tag(13, 'Important') }));

	assert.equal(markdown, '> [!important] Important\n> first\n>\n> second');
});

test('differently tagged paragraphs stay separate', async () => {
	const markdown = await render(
		para('first', { tags: tag(13, 'Important') }),
		para('second', { tags: tag(15, 'Question') }));

	assert.equal(markdown, '> [!important] Important\n> first\n\n> [!question] Question\n> second');
});

test('a tagged list item keeps its place in the list', async () => {
	const markdown = await render(
		para('one', { list: { level: 0, ordered: false } }),
		para('two', { list: { level: 0, ordered: false }, tags: tag(13, 'Important') }),
		para('three', { list: { level: 0, ordered: false } }));

	assert.equal(markdown, '- one\n- two\n- three');
});

test('a highlight carries the circle for its colour', async () => {
	assert.equal(await render(para([{ text: 'lit', highlight: '#ffff00' }])), '==🟡lit==');
	assert.equal(await render(para([{ text: 'both', highlight: '#ffff00', bold: true }])), '**==🟡both==**');
});

test('a neutral code background is not treated as semantic highlighting', async () => {
	assert.equal(await render(para([{ text: 'plain', highlight: '#f2f2f2' }])), 'plain');
});

test('a long monospace assignment paragraph becomes fenced code', async () => {
	const expression = 'df_result[\'keep\'] = (df_result[\'score\'] >= 0.95) & (~df_result[\'sample\'].isin(samples_to_remove)) & (df_result[\'count\'] > 10)';
	assert.equal(
		await render(para([{ text: expression, font: 'Consolas', highlight: '#f2f2f2' }])),
		`\`\`\`\n${expression}\n\`\`\``);
});

test('a prose explanation that starts with an assignment remains prose', async () => {
	const markdown = await render(para([{
		text: 'min_log2FoldChange = 0.5, which means that the fold change is greater than one.',
		font: 'Consolas',
	}]));
	assert.equal(markdown, 'min_log2FoldChange = 0.5, which means that the fold change is greater than one.');
});

test('an assignment explanation does not open a Python block before real code', async () => {
	const markdown = await render(
		para('min_log2FoldChange = 0.5, which means that the fold change is greater than one.'),
		para('import pandas as pd'),
		para('frame = pd.DataFrame()'));
	assert.equal(markdown, ['min_log2FoldChange = 0.5, which means that the fold change is greater than one.', '', '```python', 'import pandas as pd', 'frame = pd.DataFrame()', '```'].join('\n'));
});

test('consecutive red syntax-highlighted config paragraphs become one clean code block', async () => {
	const lines = [
		'# config file',
		'database_name =',
		'decoy_search = 1',
		'num_threads = 12',
		'# masses',
		'peptide_mass_tolerance = 10.00',
		'[ENZYME_INFO]',
		'1. Trypsin 1 KR P',
	];
	const elements = lines.flatMap((text, index) => [
		para([{ text, highlight: '#ff0000' }]),
		...(index === 3 ? [para('')] : []),
	]);

	assert.equal(await render(...elements), `\`\`\`\n${lines.slice(0, 4).join('\n')}\n\n${lines.slice(4).join('\n')}\n\`\`\``);
});

test('red config paragraphs in simple outlines still form one code block', async () => {
	const line = (text: string): Element => ({ kind: 'outline', children: [para([{ text, highlight: '#ff0000' }])] });
	const elements = [
		line('# config'), line('database_name ='), line('decoy_search = 1'), line('num_threads = 12'),
		line('# masses'), line('peptide_mass_tolerance = 10'), line('[ENZYME_INFO]'), line('1. Trypsin 1 KR P'),
	];

	assert.match(await render(...elements), /^```\n# config\n[\s\S]*\n```$/);
});

test('rendered red config lines separated by blank paragraphs are cleaned after conversion', async () => {
	const values = ['# config', 'database_name =', 'decoy_search = 1', 'num_threads = 12', '# masses', 'peptide_mass_tolerance = 10', '[ENZYME_INFO]', '1. Trypsin 1 KR P'];
	const elements: Element[] = values.map(value => ({
		kind: 'outline',
		children: [para([{ text: value, highlight: '#ff0000' }]), para('')],
	}));
	const converted = await convertPage(page(...elements), { saveAttachment: async () => null });

	assert.doesNotMatch(converted.markdown, /==🔴/);
	assert.match(converted.markdown, /^```\n# config\n[\s\S]*\n```$/);
});

test('a colour between two of them takes the nearer', async () => {
	assert.equal(await render(para([{ text: 'amber', highlight: '#ffc000' }])), '==🟠amber==');
	assert.equal(await render(para([{ text: 'lime', highlight: '#00ff00' }])), '==🟢lime==');
	assert.equal(await render(para([{ text: 'cyan', highlight: '#00ffff' }])), '==🔵cyan==');
	assert.equal(await render(para([{ text: 'magenta', highlight: '#ff00ff' }])), '==🟣magenta==');
});

test('the emphases markdown has no syntax for fall back to HTML', async () => {
	assert.equal(await render(para([{ text: 'x', superscript: true }])), '<sup>x</sup>');
	assert.equal(await render(para([{ text: 'y', subscript: true }])), '<sub>y</sub>');
	assert.equal(await render(para([{ text: 'z', underline: true }])), '<u>z</u>');
});

test('a table cell keeps the picture in it', async () => {
	const saved: string[] = [];

	const converted = await convertPage(page({
		kind: 'table',
		rows: [{ cells: [
			{ children: [{ kind: 'image', fileName: 'in-cell.png', data: new Uint8Array([1, 2, 3]) }] },
			{ children: [para('beside it')] },
		] }],
	}), { saveAttachment: async (_bytes, name) => {
		saved.push(name); return { path: name, name }; 
	} });

	assert.equal(converted.markdown.split('\n')[0], '| ![in-cell.png](Test%20image.png) | beside it |');
	assert.deepEqual(saved, ['Test image.png']);
});

test('an empty single-cell wrapper preserves its nested table content', async () => {
	const converted = await convertPage(page({
		kind: 'table',
		rows: [{ cells: [{ children: [{ kind: 'table', rows: [{ cells: [{ children: [para('inner')] }] }] }] }] }],
	}), {
		saveAttachment: async () => null,
	});

	assert.equal(converted.markdown, 'inner');
	assert.equal(converted.htmlFallbacks, 0);
});

test('a single-cell wrapper around a two-column callout keeps the callout table', async () => {
	const converted = await convertPage(page({
		kind: 'table',
		rows: [{ cells: [{ children: [{ kind: 'outline', children: [{
			kind: 'table',
			rows: [{ cells: [
				{ children: [para('Information')] },
				{ children: [para('Check the related page')] },
			] }],
		}] }] }] }],
	}), { saveAttachment: async () => null });

	assert.equal(converted.markdown, [
		'| Information | Check the related page |',
		'| --- | --- |',
	].join('\n'));
	assert.equal(converted.htmlFallbacks, 0);
});

test('a monospace single-cell table becomes a multiline code block', async () => {
	const markdown = await render({
		kind: 'table',
		rows: [{ cells: [{ children: [
			para([{ text: 'import pandas as pd', font: 'Consolas', bold: true }]),
			para([{ text: 'print(pd.__version__)', font: 'Consolas' }]),
		] }] }],
	});

	assert.equal(markdown, '```python\nimport pandas as pd\nprint(pd.__version__)\n```');
});

test('an unstyled single-cell table with high-confidence R syntax becomes code', async () => {
	const markdown = await render({
		kind: 'table',
		rows: [{ cells: [{ children: [
			para('library(tidyverse)'),
			para('input <- read.csv("input.tsv", sep="\\t")'),
			para('result <- input %>% group_by(sample) %>% summarise(count = n())'),
			para('write.csv(result, "summary.csv", row.names = FALSE)'),
		] }] }],
	});

	assert.equal(markdown, [
		'```r',
		'library(tidyverse)',
		'input <- read.csv("input.tsv", sep="\\t")',
		'result <- input %>% group_by(sample) %>% summarise(count = n())',
		'write.csv(result, "summary.csv", row.names = FALSE)',
		'```',
	].join('\n'));
});

test('a VCF table remains raw VCF instead of becoming Markdown headings', async () => {
	const markdown = await render({
		kind: 'table',
		rows: [{ cells: [{ children: [
			para('##fileformat=VCFv4.1'),
			para('##INFO=<ID=DP,Number=1,Type=Integer,Description="Read depth">'),
			para('#CHROM POS ID REF ALT QUAL FILTER INFO'),
			para('chr1 100 . A G 60 PASS DP=12'),
		] }] }],
	});

	assert.equal(markdown, [
		'```vcf',
		'##fileformat=VCFv4.1',
		'##INFO=<ID=DP,Number=1,Type=Integer,Description="Read depth">',
		'#CHROM POS ID REF ALT QUAL FILTER INFO',
		'chr1 100 . A G 60 PASS DP=12',
		'```',
	].join('\n'));
});

test('consecutive VCF paragraphs remain raw VCF instead of becoming Markdown headings', async () => {
	const markdown = await render(
		para('before'),
		para('##fileformat=VCFv4.1'),
		para('##INFO=<ID=DP,Number=1,Type=Integer,Description="Read depth">'),
		para('#CHROM POS ID REF ALT QUAL FILTER INFO'),
		para('chr1 100 . A G 60 PASS DP=12'),
	);

	assert.equal(markdown, [
		'before',
		'',
		'```vcf',
		'##fileformat=VCFv4.1',
		'',
		'##INFO=\\<ID=DP,Number=1,Type=Integer,Description="Read depth">',
		'',
		'#CHROM POS ID REF ALT QUAL FILTER INFO',
		'',
		'chr1 100 . A G 60 PASS DP=12',
		'```',
	].join('\n'));
});

test('a VCF wrapped in a non-simple outline remains raw VCF', async () => {
	const wrapped = (text: string): Element => ({ kind: 'outline', children: [para(text), para('')] });
	const converted = await convertPage(page(
		wrapped('##fileformat=VCFv4.1'),
		wrapped('##INFO=<ID=DP,Number=1,Type=Integer,Description="Read depth">'),
		wrapped('#CHROM POS ID REF ALT QUAL FILTER INFO'),
		wrapped('chr1 100 . A G 60 PASS DP=12'),
	), { saveAttachment: async () => null });

	assert.match(converted.markdown, /^```vcf\n##fileformat=VCFv4.1/m);
	assert.match(converted.markdown, /chr1 100 \. A G 60 PASS DP=12\n```$/);
});

test('a long Python dictionary remains code rather than prose', async () => {
	const value = `{'protein_description': 'ALL_1.p1', 'AA_seq': '${'MPEPTIDE'.repeat(20)}', 'source': 'analysis'}`;
	const converted = await convertPage(page(para(value)), { saveAttachment: async () => null });
	assert.equal(converted.markdown, `\`\`\`python\n${value}\n\`\`\``);
});

test('a long Python dictionary restores Markdown-escaped brackets inside code', async () => {
	const value = `{'locations': \\[1, 2, 3\\], 'AA_seq': '${'MPEPTIDE'.repeat(20)}'}`;
	const converted = await convertPage(page(para(value)), { saveAttachment: async () => null });
	assert.equal(converted.markdown, `\`\`\`python\n${value.replaceAll('\\\\', '')}\n\`\`\``);
});

test('an ordinary short dictionary-like note remains prose', async () => {
	assert.equal(await render(para("{'status': 'done'}")), "{'status': 'done'}");
});

test('a run of syntax-coloured Python paragraphs becomes runnable code', async () => {
	const markdown = await render(
		para([{ text: 'def', bold: true }, { text: ' count(items):' }]),
		para([{ text: 'for', bold: true }, { text: ' item ' }, { text: 'in', bold: true }, { text: ' items:' }]),
		para([{ text: 'return', bold: true }, { text: ' len(items)' }]));
	assert.equal(markdown, ['```python', 'def count(items):', 'for item in items:', 'return len(items)', '```'].join('\n'));
});

test('a Python expression continues an established code block', async () => {
	const markdown = await render(
		para('import pandas as pd'),
		para('frame = pd.DataFrame()'),
		para('frame["count"] += 1'),
		para('print(frame)'));
	assert.equal(markdown, ['```python', 'import pandas as pd', 'frame = pd.DataFrame()', 'frame["count"] += 1', 'print(frame)', '```'].join('\n'));
});

test('R package setup stays in an R block', async () => {
	const markdown = await render(
		para('library("MatrixEQTL")'),
		para('base.dir = find.package("MatrixEQTL")'),
		para('output_file_name = tempfile()'));
	assert.equal(markdown, ['```r', 'library("MatrixEQTL")', 'base.dir = find.package("MatrixEQTL")', 'output_file_name = tempfile()', '```'].join('\n'));
});

test('a sequence of shell commands becomes a bash block', async () => {
	const markdown = await render(
		para([{ text: 'cd ' }, { text: '/', bold: true }, { text: 'work' }]),
		para([{ text: 'gatk --java-options "-Xmx4G" HaplotypeCaller --input sample.bam' }]),
		para([{ text: 'samtools view -b sample.bam > sample.bam' }]));
	assert.equal(markdown, ['```bash', 'cd /work', 'gatk --java-options "-Xmx4G" HaplotypeCaller --input sample.bam', 'samtools view -b sample.bam > sample.bam', '```'].join('\n'));
});

test('a long shell loop with fragmented bold formatting becomes a bash block', async () => {
	const command = '**for** n **in** {1..18}**;** **do** python /work/run.py --sample ${n} --output /work/results/${n}**;** **done**';
	assert.equal(await render(para(command)), '```bash\nfor n in {1..18}; do python /work/run.py --sample ${n} --output /work/results/${n}; done\n```');
});

test('plain text does not continue an established bash block', async () => {
	const markdown = await render(para('cd /work'), para('wget https://example.test/data'), para('on another machine'));
	assert.equal(markdown, ['```bash', 'cd /work', 'wget https://example.test/data', '```', 'on another machine'].join('\n'));
});

test('a dense GFF record sequence becomes a gff block', async () => {
	const record = (feature: string, start: number) => `chr9 transdecoder ${feature} ${start} ${start + 50} . + . ID=ALL.p1.${feature}${start};Parent=ALL.p1`;
	const markdown = await render(
		para([{ text: 'chr9 transdecoder exon ' }, { text: '100', highlight: '#ffff00' }, { text: ' 150 . + . ID=ALL.p1.exon100;Parent=ALL.p1' }]),
		para(record('CDS', 200)),
		para(record('three_prime_UTR', 300)));
	assert.equal(markdown, ['```gff', record('exon', 100), record('CDS', 200), record('three_prime_UTR', 300), '```'].join('\n'));
});

test('an unstyled single-cell table with high-confidence Python syntax becomes code', async () => {
	const markdown = await render({
		kind: 'table',
		rows: [{ cells: [{ children: [
			para('import pandas as pd'),
			para('input_file = r"C:\\data\\input.tsv"'),
			para('frame = pd.read_csv(input_file, sep="\\t")'),
			para('frame.to_csv("summary.tsv", sep="\\t", index=False)'),
		] }] }],
	});

	assert.equal(markdown, [
		'```python',
		'import pandas as pd',
		'input_file = r"C:\\data\\input.tsv"',
		'frame = pd.read_csv(input_file, sep="\\t")',
		'frame.to_csv("summary.tsv", sep="\\t", index=False)',
		'```',
	].join('\n'));
});

test('ordinary prose in an unstyled single-cell table is not guessed as code', async () => {
	const markdown = await render({
		kind: 'table',
		rows: [{ cells: [{ children: [
			para('Import the data from the shared folder.'),
			para('Read the results and check every sample.'),
			para('Write a short summary for the next meeting.'),
		] }] }],
	});

	assert.equal(markdown, [
		'Import the data from the shared folder.',
		'',
		'Read the results and check every sample.',
		'',
		'Write a short summary for the next meeting.',
	].join('\n'));
});

test('a table of contents wrapper is not guessed as code from heading markers', async () => {
	const markdown = await render({
		kind: 'table',
		rows: [
			{ cells: [{ children: [para('Table of Contents [Refresh]')] }] },
			{ cells: [{ children: [para('# First section')] }] },
			{ cells: [{ children: [para('# Second section')] }] },
		],
	});

	assert.doesNotMatch(markdown, /^```/);
	assert.match(markdown, /Table of Contents/);
});

test('a line-number column is removed from a code block', async () => {
	const markdown = await render({
		kind: 'table',
		rows: [{ cells: [
			{ children: [para([{ text: '1', font: 'Consolas' }]), para([{ text: '2', font: 'Consolas' }])] },
			{ children: [para([{ text: 'library(tidyverse)', font: 'Consolas' }]), para([{ text: 'x <- 1', font: 'Consolas' }])] },
		] }],
	});

	assert.equal(markdown, '```r\nlibrary(tidyverse)\nx <- 1\n```');
});

test('trailing blank code paragraphs do not break line-number detection', async () => {
	const markdown = await render({
		kind: 'table',
		rows: [{ cells: [
			{ children: [para('1'), para('2')] },
			{ children: [para('echo first'), para('echo second'), para(''), para('')] },
		] }],
	});

	assert.equal(markdown, '```\necho first\necho second\n```');
});

test('a small line-count discrepancy keeps high-confidence numbered code', async () => {
	const markdown = await render({
		kind: 'table',
		rows: [{ cells: [
			{ children: [para('1'), para('2'), para('3')] },
			{ children: [para('import os'), para(''), para('print(os.getcwd())'), para('print("done")')] },
		] }],
	});

	assert.equal(markdown, '```python\nimport os\n\nprint(os.getcwd())\nprint("done")\n```');
});

test('embedded line breaks count as separate code lines', async () => {
	const markdown = await render({
		kind: 'table',
		rows: [{ cells: [
			{ children: [para('1\n2\n3')] },
			{ children: [para('import os\nprint(os.getcwd())\nprint("done")')] },
		] }],
	});

	assert.equal(markdown, '```python\nimport os\nprint(os.getcwd())\nprint("done")\n```');
});

test('a structural code table survives missing font metadata', async () => {
	const markdown = await render({
		kind: 'table',
		rows: [{ cells: [
			{ children: [para('')] },
			{ children: [para('cd /data/project'), para('for sample in *.fastq.gz; do'), para('\techo "$sample"'), para('done')] },
		] }],
	});

	assert.equal(markdown, '```bash\ncd /data/project\nfor sample in *.fastq.gz; do\n\techo "$sample"\ndone\n```');
});

test('a one-row number and shell-command table becomes bash code', async () => {
	const markdown = await render({
		kind: 'table',
		rows: [{ cells: [
			{ children: [para('1')] },
			{ children: [para('mapper.pl /data/reads.fa -d -p /data/genome.fa -o 48 -s reads.fa -t reads.arf')] },
		] }],
	});
	assert.equal(markdown, '```bash\nmapper.pl /data/reads.fa -d -p /data/genome.fa -o 48 -s reads.fa -t reads.arf\n```');
});

test('an absolute-path mapper command in a one-row table becomes bash code', async () => {
	const command = '/home/xcao/p/anaconda3/envs/bio/bin/mapper.pl /data/reads.fa -d -p /data/genome.fa -o 48 -s reads.fa -t reads.arf -v -j -l 18 -m';
	const markdown = await render({
		kind: 'table',
		rows: [{ cells: [{ children: [para('1'), para('')] }, { children: [para(command)] }] }],
	});
	assert.equal(markdown, `\`\`\`bash\n${command}\n\`\`\``);
});

test('a numbered GFM shell-command table is normalized to bash code', async () => {
	const command = '/home/xcao/p/anaconda3/envs/bio/bin/mapper.pl /data/reads.fa -d -p /data/genome.fa -o 48 -s reads.fa -t reads.arf -v -j -l 18 -m';
	const converted = await convertPage(page(para(`| 1 | ${command} |\n| --- | --- |`)), { saveAttachment: async () => null });
	assert.equal(converted.markdown, `\`\`\`bash\n${command}\n\`\`\``);
});

test('a real nested two-dimensional table falls back to local HTML', async () => {
	const converted = await convertPage(page({
		kind: 'table',
		rows: [{ cells: [
			{ children: [para('parent')] },
			{ children: [{ kind: 'table', rows: [{ cells: [{ children: [para('child A')] }, { children: [para('child B')] }] }] }] },
		] }],
	}), { saveAttachment: async () => null });

	assert.match(converted.markdown, /^<table>/);
	assert.match(converted.markdown, /<table>.*child A.*child B.*<\/table>/s);
	assert.equal(converted.htmlFallbacks, 1);
});

test('a deeply nested two-dimensional table stays as one local HTML structure', async () => {
	const converted = await convertPage(page({
		kind: 'table',
		rows: [
			{ cells: [
				{ children: [{ kind: 'table', rows: [
					{ cells: [{ children: [para('1')] }] },
					{ cells: [{ children: [para('2')] }] },
				] }] },
				{ children: [para('3')] },
			] },
			{ cells: [{ children: [para('4')] }, { children: [para('5')] }] },
			{ cells: [
				{ children: [para('6')] },
				{ children: [para('7'), { kind: 'table', rows: [
					{ cells: [{ children: [para('ab')] }, { children: [para('cd')] }] },
					{ cells: [
						{ children: [para('ef')] },
						{ children: [para('gh'), { kind: 'table', rows: [{ cells: [
							{ children: [para('OK')] },
							{ children: [para('OOKK')] },
						] }] }] },
					] },
				] }] },
			] },
			{ cells: [{ children: [para('')] }, { children: [para('')] }] },
			{ cells: [{ children: [para('888')] }, { children: [para('999')] }] },
		],
	}), { saveAttachment: async () => null });

	assert.equal(converted.htmlFallbacks, 1);
	assert.equal((converted.markdown.match(/<table>/g) ?? []).length, 4);
	for (const value of ['1', '2', '3', '4', '5', '6', '7', 'ab', 'cd', 'ef', 'gh', 'OK', 'OOKK', '888', '999']) {
		assert.match(converted.markdown, new RegExp(`>${value}<`));
	}
});

test('HTML table fallback turns OneNote vertical tabs into line breaks', async () => {
	const markdown = await render({
		kind: 'table',
		rows: [{ cells: [
			{ children: [{ kind: 'table', rows: [{ cells: [
				{ children: [para([{ text: 'before\u000bafter' }])] },
				{ children: [para([{ text: 'right' }])] },
			] }] }] },
			{ children: [para([{ text: 'outer' }])] },
		] }],
	});
	assert.match(markdown, /before<br>after/);
	assert.doesNotMatch(markdown, /\u000b/);
});

test('text that looks like markdown is not read as markdown', async () => {
	assert.equal(await render(para('# ordinary text')), '\\# ordinary text');
	assert.equal(await render(para('- not a list')), '\\- not a list');
	assert.equal(await render(para('> not a quote')), '\\> not a quote');
	assert.equal(await render(para('1. not numbered')), '\\1. not numbered');
	assert.equal(await render(para('see [1](x)')), 'see \\[1\\](x)');
	assert.equal(await render(para('use `code`')), 'use \\`code\\`');
	assert.equal(await render(para('a <b> tag')), 'a \\<b> tag');
});

test('ordinary prose is left unmarked', async () => {
	assert.equal(await render(para('a * b * c')), 'a * b * c');
	assert.equal(await render(para('file_name_here')), 'file_name_here');
	assert.equal(await render(para('2 - 1 = 1')), '2 - 1 = 1');
});

test('an escaped line keeps the formatting the importer added', async () => {
	assert.equal(
		await render(para('# text', { list: { level: 0, ordered: false } })),
		'- \\# text');
});

test('a link out of the notebook is left exactly as it was', async () => {
	assert.equal(
		await render(para([{ text: 'a site', hyperlinkUrl: 'https://example.com/a b' }])),
		'[a site](https://example.com/a%20b)');
	assert.equal(
		await render(para([{ text: 'a file', hyperlinkUrl: 'file:///C:/notes.txt' }])),
		'[a file](file:///C:/notes.txt)');
});

test('a leaked Office hyperlink field becomes an ordinary external Markdown link', async () => {
	const converted = await render(para('\uf7dfHYPERLINK "https://example.com/a b"visible text'));
	assert.equal(converted, '[visible text](https://example.com/a%20b)');
});

test('a private-use variant of a leaked Office hyperlink field is removed too', async () => {
	const converted = await render(para('\ufdefHYPERLINK "https://example.com/a b"visible text'));
	assert.equal(converted, '[visible text](https://example.com/a%20b)');
});

test('a leaked OneNote hyperlink field keeps only its visible text', async () => {
	const converted = await render(para('\uf7dfHYPERLINK "onenote:#Other%20page&section-id={section}"visible text'));
	assert.equal(converted, 'visible text');
});

test('literal formatted XML becomes a local XML code block', async () => {
	const markdown = await render(
		para('**\\<?xml** version="1.0"**?>**'),
		para('**\\<settings>**'),
		para('**\\<value>**x**\\</value>**'),
		para('**\\</settings>**'),
	);
	assert.equal(markdown, '```\n<?xml version="1.0"?>\n\n<settings>\n\n<value>x</value>\n\n</settings>\n```');
});

test('a OneNote link with no page in it keeps only its visible text', async () => {
	const url = 'onenote:///C:/Notebook/Section.one';
	assert.equal(await render(para([{ text: 'the section', hyperlinkUrl: url }])), 'the section');
});

test('a OneMore navigation link keeps only its visible text', async () => {
	assert.equal(await render(para([{ text: 'Refresh', hyperlinkUrl: 'onemore://InsertTocCommand/refresh/links/style0' }])), 'Refresh');
});

test('an unresolvable page name does not retain a OneNote URL', async () => {
	assert.equal(
		await render(para([{
			text: 'page',
			hyperlinkUrl: 'onenote:https://d.docs.live.net/x/Nb/S.one#Q4%20review&section-id={a}&page-id={b}&end',
		}])),
		'page');
});

test('the importer decides what an internal link points at', async () => {
	const converted = await convertPage(page(para([{
		text: 'link',
		hyperlinkUrl: 'onenote:///S.one#Notes: Q4/Q1&section-id={a}',
	}])), {
		saveAttachment: async () => null,
		resolveInternalLink: title => title.replace(/[:/]/g, '-'),
	});

	assert.equal(converted.markdown, '[link](Notes-%20Q4-Q1)');
});

test('an attachment is named for the note, not the page title it came from', async () => {
	// "10/18" is a legal page title and an illegal file name; the importer
	// says what the note ended up called.
	const converted = await convertPage(page({ kind: 'image', extension: '.png', data: new Uint8Array([1]) }), {
		noteName: '10-18',
		saveAttachment: async (_bytes, name) => ({ path: name, name }),
	});

	assert.equal(converted.attachments[0].name, '10-18 image.png');
});
