# Release notes

## v1.0.20 — Public release hygiene and release automation

### Highlights

- Removed legacy chronological change-history files from the private snapshot and public artefact staging to reduce accidental release of project-history detail.
- Added this stable `assets/release-notes.md` file for GitHub release creation. The filename intentionally has no version number; git tags and release metadata carry the version.
- Updated the userscript README generation so the displayed version and description are refreshed on every public build.
- Updated public build/sync scripts so release notes are copied to the public userscript repository.
- Added release automation for build, sync, git push, and GitHub release creation.

### Userscript artefact

- File: `m365-copilot-export.js`
- Version: `1.0.19`
- Description: Export the current Microsoft 365 Copilot Chat conversation to readable Markdown and raw JSON Markdown files.

### Publication notes

- GitHub release notes should use this file via `gh release create ... --notes-file assets/release-notes.md`.
- GreasyFork publication can remain manual or sync automatically from the public userscript repository.


## v1.0.20 — Release automation argument fix

### Highlights

- Repaired `tools/publish-release.ps1` so Bun, Git, and GitHub CLI commands receive their subcommands and arguments reliably.
- Replaced positional `Run bun @(...)` style calls with named `Invoke-External -FilePath ... -ArgumentList ...` calls.
- Kept `-DryRun` support for previewing the release flow before pushing or creating releases.

### Userscript artefact

- Version: `1.0.19`


## v1.0.20 — Release metadata template fix

### Highlights

- Fixed tools/update-release-metadata.mjs so build:all no longer fails on nested template-literal backticks in generated README content.
- Kept release automation argument handling from the previous release.
