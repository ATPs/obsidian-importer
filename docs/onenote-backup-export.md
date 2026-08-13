# OneNote Backup Export Workflow

This guide covers the local Node tools added for converting OneNote backup
directories outside Obsidian. It is the required workflow for multi-version
backup exports. The source backups are read-only.

## Tools

### `scripts/convert-onenote-local.ts`

Runs the existing OneNote importer against one `.one`, `.onepkg`, `.onex` file
or a directory of those files, using the test vault shim. It writes a fresh
portable Markdown tree and `_conversion-report.json`.

```powershell
.\node_modules\.bin\tsx.cmd --disable-warning=ExperimentalWarning --tsconfig tsconfig.test.json `
  scripts\convert-onenote-local.ts `
  "H:\path\to\section-or-directory" `
  "H:\path\to\fresh-output"
```

The destination must not exist. Attachments are written to the `attachments`
folder beside each note's parent path.

### `scripts/export-onenote-backups.ts`

Builds one final Markdown notebook per input backup directory. It converts all
versions into a private staging area, chooses the newest current page text by
OneNote page identity and timestamp, and recovers an older attachment only
when page ownership, filename/type, bytes, and label correspondence are
defensible. It writes `_merge-report.json` at the output root.

```powershell
.\node_modules\.bin\tsx.cmd --disable-warning=ExperimentalWarning --tsconfig tsconfig.test.json `
  scripts\export-onenote-backups.ts `
  "H:\x\temp\md-candidate-N" `
  "H:\x\temp\HumanNovelProteins" `
  "H:\x\temp\MouseOmics" `
  "H:\x\temp\NeuropathicPain2020" `
  "H:\x\temp\Rutgers"
```

Do not use `--reuse-staging` after changing conversion code. A full candidate
must be rebuilt from the original backup files. Immediately move the generated
`_staging` directory outside the candidate root before auditing:

```powershell
Move-Item "H:\x\temp\md-candidate-N\_staging" "H:\x\temp\md-candidate-N-staging"
```

The candidate root must then contain only the four notebook directories and
`_merge-report.json`.

### `scripts/audit-onenote-output.ts`

Validates a final candidate or published tree. It checks page count and unique
OneNote IDs, generated Markdown structure, local page links, attachment
manifests and byte hashes, recovery/missing-asset metadata, HTML fallbacks,
source hashes, and unexpected root entries. It writes JSON to stdout and exits
nonzero when a required invariant fails.

```powershell
.\node_modules\.bin\tsx.cmd --disable-warning=ExperimentalWarning --tsconfig tsconfig.test.json `
  scripts\audit-onenote-output.ts `
  "H:\x\temp\md-candidate-N" `
  "H:\x\temp\HumanNovelProteins" `
  "H:\x\temp\MouseOmics" `
  "H:\x\temp\NeuropathicPain2020" `
  "H:\x\temp\Rutgers" `
  --previous="H:\x\temp\md-improvement-notes\final-audit.json" `
  | Set-Content -Encoding utf8 "H:\x\temp\md-improvement-notes\candidate-N-audit.json"
```

`--previous` is an audit JSON file, not an output directory. Use it only for a
numeric comparison with the previous published audit.

## Required Candidate Checks

Before publishing, require an audit `failureCount` of `0` and inspect:

```powershell
rg -l -i 'onenote:|onemore:|HYPERLINK\s+"|\*\*\\?<\?xml' "H:\x\temp\md-candidate-N" --glob '*.md'
```

The search must return no files. Confirm the audit has all expected current
pages and unique IDs; zero preservation headings, XML fences, archive files,
retained OneNote links, invalid local links, and unclosed fences; a one-to-one
manifest/body/disk attachment count; reviewed localized HTML fallback tables;
and structured missing-asset entries only where recovery is unsafe.

After code changes, run:

```powershell
.\node_modules\.bin\tsx.cmd --disable-warning=ExperimentalWarning --tsconfig tsconfig.test.json --test "tests/onenote-file/**/*.test.ts"
.\node_modules\.bin\tsc.cmd --noEmit --skipLibCheck
.\node_modules\.bin\eslint.cmd src
node esbuild.config.mjs production
git diff --check
```

Injected asset-write failure stack traces are expected recovery tests; the test
command itself must still exit `0`.

## Link, Attachment, and Fallback Policy

- A uniquely resolved internal page link becomes a local relative Markdown link.
- Cross-notebook, unresolved, ambiguous, OneNote, and OneMore navigation links
  retain only visible text. Never retain a OneNote-specific URL.
- Preserve newest page text. Recover an older attachment only under the strict
  matcher; do not attach similarly named files speculatively.
- Use Markdown for ordinary content and tables. A nested one-column structure
  becomes a hierarchy/list where that preserves its relationship. Use local
  HTML only for a genuine nested two-dimensional table; inspect each fallback.
- Do not emit `_OneNote archive.md`, a preservation footer, whole-page XML, or
  XML-language fences for this backup-export workflow. Literal XML source text
  may be retained in an ordinary fenced code block.

## Page-Level Review Records

Keep one CSV per notebook under `H:\x\temp\md-improvement-notes`:

```text
path,baseline_sha256,current_sha256,status,note
```

Use `improved`, `no_change`, or `needs_manual`; do not leave completed reviews
as `pending`. Each note needs page-specific evidence. After publishing a
regenerated tree, recompute every `current_sha256` against published Markdown
while retaining the review status and evidence notes.

## Publishing

1. Copy the verified candidate to a dated snapshot inside
   `H:\x\temp\md-improvement-notes`.
2. Rename the existing `H:\x\temp\md` to a timestamped `md-backup-*` directory.
3. Rename the verified candidate to `H:\x\temp\md`.
4. Audit the published path and update
   `H:\x\temp\md-improvement-notes\final-audit.json`.
5. Verify the published root contains only four notebooks and `_merge-report.json`.

Do not delete older candidate, staging, or backup directories without explicit
user approval. They are recoverable history, not build residue.
