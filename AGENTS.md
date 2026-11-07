# AGENTS.md

Instructions for coding agents working in this repository. The project context is in the
[README](README.md); the standards, explained for people, are in [CONTRIBUTING](CONTRIBUTING.md).
This file is the short, operational version.

## The project

Offline desktop editor for documents and spreadsheets. Electron 43 + React 19 + strict
TypeScript, with a .NET 10 format service (`sidecar/`) that reads and writes DOCX, ODT and XLSX.
The core promise: **saving rewrites only what was edited**; the rest of the package comes back
byte for byte.

## Commands

| Command | When |
| --- | --- |
| `npm run typecheck` | after any TypeScript change |
| `npm run lint` | same; enforces the layer boundaries and the function limits |
| `npx vitest run <file>` | unit tests for what you touched |
| `npm run sidecar:test` | changes in `sidecar/` |
| `npm run sidecar:build` | before tests that talk to the published sidecar, and before e2e |
| `npm run e2e` | build + sidecar + Playwright; only when the change touches the interface |
| `npm run format:check` / `npm run format` | check and fix formatting |
| `npm run verify` | everything CI runs, except formatting |

Test only what the change touched. Running the whole suite at every step is waste.

**A bug fix starts with the test.** Before touching the code, write a test, in the right layer,
that reproduces the bug, and run it to watch it fail for the bug's reason. Only then fix it, and
run it again to watch it pass. The test goes in the same commit as the `fix`.

## Rules the code enforces

- **Boundaries:** `src/services/` and `src/shared/` do not import `electron`, `react`, `node:*`,
  `@main/*` or `@renderer/*`. The renderer does not import `electron`, `node:*` or `@main/*`; it
  talks to main only through `window.api`. The preload imports only from `@shared`.
- **Function limits** (in `src/`, outside tests): `complexity` 15, `max-lines-per-function` 80,
  `max-depth` 4, `max-params` 5. No function exceeds them, and there is no exceptions list: above
  the limit, split the function.
- **`eslint-disable`** only on the line, with the reason written above it, and only when the
  cause is external (for example, a ProseMirror callback signature).
- **IPC:** every channel has a Zod schema in `src/shared/`.
- **Sidecar protocol:** `src/main/sidecar/protocol.ts` and
  `sidecar/src/Librevia.Format/Protocol/Frame.cs` change together.
- **Own formats:** if what `.sdoc` or `.ssheet` stores changes, bump `SDOC_VERSION` or
  `SSHEET_VERSION` and keep reading the old versions.
- **Loss warnings and sidecar errors:** the sidecar sends the key, never the sentence. Use the
  constants in `sidecar/src/Librevia.Format/Docx/Inventory.cs`; a new key goes in
  `src/shared/i18n/catalog/sidecar.ts`, with `pt` and `en`.
- **Interface text:** every visible sentence goes in the `src/shared/i18n/catalog/` catalog, with
  `pt` and `en`. `npm run i18n:check` flags loose sentences in the code.
- **Network:** the application does not access the internet. Do not add anything that fetches
  from outside.
- **MCP server:** the guarantees are in [SECURITY.md](SECURITY.md#in-scope) and do not
  change unless the owner asks. Day to day: a new tool goes only in `document-tools.ts` or
  `sheet-tools.ts`, does not read files, does not access the network and does not execute
  anything; every write becomes a tracked change, a comment or a proposal on the sheet, never a
  direct write.
- **A new dependency** needs a permissive license; the license gate rejects the rest.

## Style

- The principles (Clean Code, Clean Architecture, SOLID) are in
  [CONTRIBUTING](CONTRIBUTING.md#principles); in practice, business rules live in `src/services/`,
  without Electron or React.
- Identifiers, the rare comments and tests (names, descriptions and comments) in **English**.
  User-facing messages in **Portuguese**.
- **Avoid comments.** The code should explain itself through names, small functions and named
  constants; the why goes in the test name and the commit. A comment only for what the code
  cannot say (an external constraint, counterintuitive Word or Excel behavior, the reason for an
  `eslint-disable`), and always saying why, never what.
- When touching code with a comment that only repeats it, replace the comment with a better name.
- **Search before writing a helper** (`git grep` for what it does); the common ones are in the
  table in [Search before you write](CONTRIBUTING.md#search-before-you-write). No local copies; on
  the second occurrence, extract it to the right layer and delete the copies in the same commit.
- README and MANUAL: simple, exact language, short sentences, concrete examples, measured numbers.
- `README.md`, `MANUAL.md`, `CONTRIBUTING.md` and `SECURITY.md` are the source of truth; each
  has a `.pt-BR.md` translation. Changing one, update its translation in the same commit.

## Commits

- [Conventional Commits](https://www.conventionalcommits.org) in English:
  `type(scope): description`. Scopes come from the code (`docx`, `xlsx`, `pagination`, `lists`,
  `comments`, `revisions`, `notes`, `equations`, `sidecar`, `ui`…).
- **Short body or none:** one or two sentences with the why, only when the title is not enough.
- **No attribution footer** — no agent `Co-Authored-By`, no session link, no "Generated with".
  This also applies to pull request descriptions.
- **Author and committer are always the repository owner**, from the git configuration. Never use
  `--author` or `GIT_AUTHOR_*` / `GIT_COMMITTER_*` variables.
- Do not commit or push unless the owner asks.

## What stays out of the repository

- The real documents used to calibrate the project (the "corpus"). Tests that depend on them are
  skipped without `LIBREVIA_CORPUS_DOC`.
- Visual comparison and measurement scripts made during a task: they stay outside the tree.
- The `docs/` folder, which git ignores. Do not mention it in published documentation.
- Sidecar binaries (`resources/sidecar/`, `bin/`, `obj/`) and build outputs.
