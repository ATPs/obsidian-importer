![Obsidian Importer screenshot](/images/social.png)

This Obsidian plugin allows you to import notes from other apps and file formats into your Obsidian vault. Notes are converted to plain text Markdown files.

## Get started

Install Importer in Obsidian → Community Plugins.

Import guides are hosted on the [official Obsidian Help site](https://obsidian.md/help/import). You can help contribute to the guides on the [obsidian-help](https://github.com/obsidianmd/obsidian-help) repo.

- [Import from Airtable](https://obsidian.md/help/import/airtable)
- [Import from Apple Notes](https://obsidian.md/help/import/apple-notes)
- [Import from Bear](https://obsidian.md/help/import/bear)
- [Import from CSV files](https://obsidian.md/help/import/csv)
- [Import from Evernote](https://obsidian.md/help/import/evernote)
- [Import from Google Keep](https://obsidian.md/help/import/google-keep)
- [Import from Microsoft OneNote](https://obsidian.md/help/import/onenote)
- [Import from Notion](https://obsidian.md/help/import/notion)
- [Import from Roam Research](https://obsidian.md/help/import/roam)
- [Import from HTML files](https://obsidian.md/help/import/html)
- [Import from Markdown files](https://obsidian.md/help/import/markdown)
- [Import from Textbundle files](https://obsidian.md/help/import/textbundle) (.textbundle, .textpack)
- Import from Apple Journal (HTML export)
- Import from Tomboy/Gnote (.note)

## OneNote local files and backups

The Obsidian importer supports local `.one`, `.onepkg`, and `.onex` files.
For an ordinary interactive import, use **Microsoft OneNote** in Obsidian's
Importer and select the local file. For repeatable, file-system based exports
or multiple OneNote backup versions, use the local tools below from a checkout
of this repository. They read the sources; always export to a new directory.

### Export one file or one section backup

Use this when there is one source file, or when you deliberately want each
source file converted independently:

```powershell
.\node_modules\.bin\tsx.cmd --disable-warning=ExperimentalWarning --tsconfig tsconfig.test.json `
  scripts\convert-onenote-local.ts `
  "H:\OneNote backups\Research.one" `
  "H:\exports\Research-markdown"
```

The destination must not already exist. The exporter writes Markdown, adjacent
`attachments` directories, and `_conversion-report.json`.

### Merge multiple backups of one notebook

When a notebook directory contains multiple saved versions of its sections,
use the backup merger. It converts every version into private staging, then
publishes one Markdown tree for each input notebook directory:

```powershell
.\node_modules\.bin\tsx.cmd --disable-warning=ExperimentalWarning --tsconfig tsconfig.test.json `
  scripts\export-onenote-backups.ts `
  "H:\exports\md-candidate-1" `
  "H:\OneNote backups\Research" `
  "H:\OneNote backups\Lab notes"
```

For each section, current pages come only from its newest backup version. For
the same page in multiple versions, its `onenote-updated` timestamp chooses
the Markdown source; the source file modification time breaks a tie or covers
a missing timestamp. Pages absent from the newest section version are treated
as deleted and are not restored from an older backup.

Older versions are used only to recover an attachment that the chosen current
page explicitly reports as missing. Recovery requires a unique, compatible
match for the original attachment identity (source name, occurrence number,
and image-versus-file type), followed by byte-length and SHA-256 verification.
The current page's Markdown is never replaced with older page text. Ambiguous
or unavailable attachments remain recorded as missing instead of being guessed.

### Batch-export a set of notebooks

When one parent directory contains one subdirectory per notebook, use the
Windows runner. It discovers direct `.one`, `.onepkg`, and `.onex` files in
each notebook directory, exports every notebook, moves staging out of the
candidate, and runs the audit. It never overwrites or publishes an existing
export.

```powershell
# Inspect the exact notebooks and source files first.
.\scripts\export-onenote-backup-set.ps1 `
  "H:\OneNote backups" `
  "H:\exports\md-candidate-1" `
  -PlanOnly

# Create and audit a new candidate.
.\scripts\export-onenote-backup-set.ps1 `
  "H:\OneNote backups" `
  "H:\exports\md-candidate-1" `
  -AuditReport "H:\exports\candidate-1-audit.json"
```

### Images, EMF, and verification

Images remain images; the exporter does not add OCR or handwriting-recognition
text. EMF figures are rendered to PNG at 300 DPI and the original EMF sidecar
is not retained. Oversized renders preserve their aspect ratio and are reduced
only when both dimensions exceed 2000 pixels, so the shorter edge is at most
2000 pixels.

The batch runner audits automatically. To audit a manually generated candidate:

```powershell
.\node_modules\.bin\tsx.cmd --disable-warning=ExperimentalWarning --tsconfig tsconfig.test.json `
  scripts\audit-onenote-output.ts `
  "H:\exports\md-candidate-1" `
  "H:\OneNote backups\Research" `
  | Set-Content -Encoding utf8 "H:\exports\candidate-1-audit.json"
```

Require `failureCount: 0` before publishing. The audit checks page IDs, local
links, manifests, attachment hashes, missing/recovered attachment records, and
the output tree. For the full workflow and safe publication procedure, see
[the OneNote backup export guide](docs/onenote-backup-export.md).

## Developers

```bash
npm install
npm run dev      # build, copy into a vault, reload the plugin
npm test         # convert every fixture and compare against its recording
npm run build    # typecheck and build for release
```

Set `OBSIDIAN_PATH` in `.env` to the plugin folder `npm run dev` should copy into, relative to your home directory:

```
OBSIDIAN_PATH='/Documents/MyVault/.obsidian/plugins'
```

### Tests

Each importer is tested by converting a real file and comparing the result with an output recorded beside it:

```
tests/notion/notion-testspace.zip           a fixture
tests/notion/expected/notion-testspace/…    what converting it produces
tests/notion/local/                         gitignored, for a file that cannot be committed
```

Run one importer's tests while working on it:

```bash
npm test -- notion
```

To record a new fixture's output, or update one after an intended change:

```bash
UPDATE_EXPECTED=1 npm test -- notion
```

That writes the output and then fails on purpose. Read what it wrote — `git diff` if it already existed — before committing it. A recording nobody reads is not a test.

Debugging an issue someone reported? Drop their export in `tests/<importer>/local/`. It is gitignored, and so is the output recorded next to it, so you can work against a real file without committing it.

### Testing against a live API

Airtable and Notion's API importers have no export file to use as a fixture, so their fixtures are saved API responses. Those go stale quietly, so each has a check that asks the real API whether its responses still have the shape the fixture assumes. They skip unless a token is set in `.env`:

```
AIRTABLE_TOKEN=pat...
NOTION_TOKEN=ntn_...
```

These only read.

### Testing against Obsidian itself

`npm test` runs the conversions outside Obsidian, against a small stand-in for its API in `tests/shims/`. To check that stand-in still agrees with the app:

```bash
npm run e2e
```

This imports fixtures through the running app — its `htmlToMarkdown`, its vault, its link settings — and compares what lands in the vault with what `npm test` recorded. It needs the [Obsidian CLI](https://obsidian.md/help/cli) and a build of your working copy installed in the active vault. It writes one folder and deletes it afterwards.

## Contributing

Importer is a community-led project. You can explore pull requests and see the credits below for reference. The Obsidian team is not actively working on adding new import capabilities, but we welcome pull requests for new formats and improvements.

Is a format missing? You can help! See our [Contribution guidelines](/CONTRIBUTING.md).

Some issues have been [tagged with #bounty](https://github.com/obsidianmd/obsidian-importer/labels/bounty).

## Credits

This plugin relies on important contributions:

- [@akosbalasko](https://github.com/akosbalasko) for Evernote import via [Yarle](https://github.com/akosbalasko/yarle) (MIT)
- [@daledesilva](https://github.com/daledesilva) for Google Keep import
- [@arthurtyukayev](https://github.com/arthurtyukayev) for Bear import
- [@xheldon](https://github.com/Xheldon) for Notion API import and Airtable import
- [@joshuatazrein](https://github.com/joshuatazrein) for Notion file-based import
- [@polyipseity](https://github.com/polyipseity) for HTML attachments
- [@8bitgentleman](https://github.com/8bitgentleman) for Roam import
- [@p3rid0t](https://github.com/p3rid0t) for Microsoft OneNote import
- [@mirnovov](https://github.com/mirnovov) for Apple Notes import
- [@wzs](https://github.com/wzs) for Apple Journal import
