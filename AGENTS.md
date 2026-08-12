# Repository Agent Guide

Read `CLAUDE.md` before changing code. It contains the repository-wide architecture, test, localization, and Obsidian review rules. This file adds rules for the local OneNote importer and its information-preservation guarantees.

## Primary requirement

For `.one`, `.onepkg`, and `.onex` imports, preserving information is more important than producing attractive Markdown. Never silently discard a page, object, property, attachment, title, path, link target, or parser failure merely because it cannot be represented cleanly.

- Produce one Markdown note for each current OneNote page.
- Put deleted pages in `_deleted` and conflict pages in `_conflicts`.
- Put section-level and otherwise unowned records in `_OneNote archive.md`.
- Historical revisions are intentionally not imported as current pages. Do not accidentally emit them as duplicates.
- Render supported content as Markdown. Preserve unsupported or partially converted data in a normal fenced `xml` block.
- A malformed object must not abort unrelated pages or source files. Report the failure and continue wherever the remaining boundaries are trustworthy.
- Do not claim that missing bytes were preserved. When a referenced payload is absent from the source, retain its name, IDs, properties, and reference, and emit an explicit `ONENOTE_ASSET_DATA_MISSING` preservation record.

## Preservation format

`src/formats/onenote-file/preservation.ts` is the canonical fallback representation.

- XML must be well-formed and must XML-escape all values. Do not build unescaped attributes or text from source data.
- Exact binary property values up to and including 64 KiB are Base64-inline in XML.
- Larger opaque values and attachments are sidecar files. Record their relative path, byte length, and SHA-256 in both preservation data and the owning note's asset manifest.
- Never replace raw bytes with a decoded or normalized approximation. A friendly rendering may be added, but exact preservation remains required.
- Every emitted sidecar needs an owner. Page data belongs to that page; section/unowned data belongs to `_OneNote archive.md`.
- Reimport may delete a stale generated sidecar only when it is demonstrably unchanged and unshared. Retain user-modified or shared files.
- Guard recursive property/object traversal against cycles and unreasonable nesting without dropping the blocked branch: emit a preservation/error record identifying it.

Imported-note frontmatter is part of the recovery format. Keep source path and SHA-256, OneNote IDs, original title, timestamps, page status, completeness state, asset manifest, and these filename fields:

- `onenote-original-filename`
- `onenote-full-output-filename`
- `onenote-output-filename`

When adding a lossy conversion, add enough typed preservation metadata to reconstruct or independently inspect the original value. Do not put raw untrusted source strings into YAML by hand; use the repository serializer.

## Names and Windows paths

Windows is a first-class development and target platform. Do not require Linux or WSL for this importer.

- Treat path components as untrusted. Handle reserved characters, control characters, trailing spaces/dots, `.` and `..`, device names such as `CON`/`NUL`/`COM1`, empty titles, and separators embedded in titles.
- Account for case-insensitive collisions and Unicode spelling/normalization collisions. Never let one output overwrite another.
- Keep names readable and do not truncate merely to meet an arbitrary short limit. Truncate only when the actual destination path requires it.
- If sanitization, an empty title, truncation, or a collision makes a name lossy, add a stable short page-ID prefix. Stability across repeated imports matters.
- Store the complete intended output filename in frontmatter even when the physical filename is shortened.
- Compute length using the final parent path, extension, conflict/deleted directory, and attachment directory. A component that fits alone can still exceed the Windows full-path limit.
- All resolved output paths must be checked to remain inside the selected destination.

## Links

Index all pages before finalizing Markdown links.

- One matching target: rewrite the OneNote URL to the imported note.
- Multiple matching targets: emit links to every candidate, numbered, with a short Markdown explanation of the ambiguity.
- No matching target: retain the original `onenote:` URL and explain that the target was not imported/resolved.
- Never choose an arbitrary duplicate target, and never erase the original target information.

## Errors and retries

Automatic retry is for transient operations only. Use `retryTransient` and exponential backoff for Windows sharing/locking and temporary I/O failures such as `EBUSY`, `EAGAIN`, `EMFILE`, `ENFILE`, `EPERM`, and timeouts.

- Do not retry deterministic parse errors, unsupported formats, corrupt structures, checksum failures, programmer errors, or validation failures.
- A retry must repeat an idempotent operation. It must not create duplicate notes or attachments.
- Exhausted retries must be visible in the import report and must not stop unrelated work.
- Preserve stable error codes; tests and audit tooling use them.

## Parser cautions

- Keep exact object IDs and raw property IDs. Masked property IDs are useful for interpretation but are not a replacement for raw IDs.
- Distinguish an absent payload from a parser lookup bug by comparing the referenced FileData GUID with all payload GUIDs in that source.
- `.onepkg` is a Cabinet archive. Validate paths, signatures, checksums, entry boundaries, declared/expanded sizes, and decompression behavior before writing entries.
- Do not introduce low arbitrary size limits as a substitute for safe parsing. At the same time, preserve decompression-bomb and path-traversal protections.
- The current OneStore reader materializes a selected source as one `Uint8Array`. It is not constant-memory streaming and very large files remain limited by host memory. Do not describe it as streaming unless this architecture changes.
- Avoid speculative recovery that can associate bytes with the wrong page. If ownership is uncertain, archive the exact data with an explanation.

## Relevant code

- `src/formats/onenote-file.ts`: vault-facing import, path planning, frontmatter, link finalization, and archives.
- `src/formats/onenote-file/convert.ts`: semantic conversion and attachment/preservation decisions.
- `src/formats/onenote-file/preservation.ts`: exact XML/Base64/sidecar fallback.
- `src/formats/onenote-file/retry.ts`: transient retry classifier and backoff.
- `src/formats/onenote-file/semantic/`: object-space interpretation.
- `src/formats/onenote-file/onestore/`: low-level OneStore parsing.
- `src/formats/onenote-file/cabinet/`: `.onepkg` Cabinet parsing.
- `tests/onenote-file/`: focused regression tests.

Keep the conversion seam described in `CLAUDE.md`: Node APIs go through `src/filesystem.ts`; vault operations remain in the importer; conversion dependencies are passed as callbacks.

## Required verification

Use the repository's pinned lockfile. On this Windows checkout, commands can be run with pnpm 10.20.0:

```powershell
npx --yes pnpm@10.20.0 test -- onenote-file
npx --yes pnpm@10.20.0 run typecheck
npx --yes pnpm@10.20.0 run lint:check
npx --yes pnpm@10.20.0 run build
git diff --check
```

Add a focused regression test for every fixed data-loss case. Assert content and recovery metadata, not only note counts. Important coverage includes empty/invalid/long/duplicate names, ambiguous and unresolved links, deleted/conflict pages, unsupported properties, small and large binary values, missing payloads, XML escaping, transient retry, Cabinet corruption, and reimport cleanup.

The full test runner may expose a Windows dynamic-import path problem in `tests/loadable/importers.test.ts`. Do not treat that known infrastructure failure as a OneNote regression, but report it exactly and still run the focused suite. Do not suppress a new failure because it occurs on Windows.

## Real-file acceptance

For a user-provided source directory, use a fresh destination and never delete or overwrite the source. `scripts/convert-onenote-local.ts` currently accepts a directory containing supported files, not a single file, and refuses an existing destination:

```powershell
.\node_modules\.bin\tsx --tsconfig tsconfig.test.json scripts\convert-onenote-local.ts `
  "H:\path\to\source" `
  "H:\path\to\new-output"

.\node_modules\.bin\tsx --tsconfig tsconfig.test.json scripts\audit-onenote-output.ts `
  "H:\path\to\source" `
  "H:\path\to\new-output"
```

Before conversion, verify that the source exists and list the selected `.one`, `.onepkg`, and `.onex` files. If a requested destination exists, do not remove it silently; use a user-approved fresh path or a recoverable timestamped backup. After conversion:

- Require zero unexpected failures and inspect every skipped item.
- Validate every preservation XML block with an XML parser.
- Verify every manifest sidecar exists and matches its recorded length and SHA-256.
- Verify each imported note records a source file whose SHA-256 matches the input.
- Compare note, asset, archive, missing-payload, and source counts with the conversion report.
- Inspect representative Markdown for readable output, ambiguous-link explanations, filename metadata, and preservation blocks.
- Treat an explicit missing-payload record as source incompleteness only after GUID-level investigation; it is not automatically a successful conversion.

Real notebooks can contain private data. Do not commit them, paste their content into tests/logs, or copy them elsewhere. Create minimal anonymized fixtures for regressions.

## Build and Obsidian deployment

`pnpm run build` generates the root `main.js`. The installed plugin is a separate copy, commonly under `<vault>/.obsidian/plugins/obsidian-importer/`.

- Building does not deploy the plugin.
- Compare hashes before copying `main.js`, `manifest.json`, or `styles.css`; copy only changed build artifacts.
- Never overwrite `data.json`, which contains local plugin settings.
- Do not write into a user's vault or installed-plugin directory unless the user explicitly asks for deployment.
- After deployment, reload the plugin or restart Obsidian before testing.

If `NODE_TLS_REJECT_UNAUTHORIZED=0` is present, warn that TLS certificate verification is disabled. Do not silently change the user's environment.
