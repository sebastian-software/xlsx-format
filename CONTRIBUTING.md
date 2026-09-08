# Contributing to xlsx-format

Thanks for your interest in contributing! Here's how to get started.

## Setup

```bash
git clone https://github.com/sebastian-software/xlsx-format.git
cd xlsx-format
pnpm install
```

## Development Workflow

```bash
pnpm run check       # TypeScript type checking
pnpm run lint        # ESLint
pnpm run format      # Oxfmt (auto-fix)
pnpm test            # Run tests once
pnpm run test:watch  # Run tests in watch mode
pnpm run build       # Build ESM + CJS bundles
pnpm run verify      # Run the full local quality gate
```

## Making Changes

1. Fork the repository and create a branch from `main`.
2. Write your code. Follow the existing style -- Oxfmt and ESLint enforce most of it.
3. Add or update tests for any changed behavior.
4. Make sure all checks pass: `pnpm run verify`
5. Use [Conventional Commits](https://www.conventionalcommits.org/) for your commit messages (e.g. `feat: add X`, `fix: handle Y`). Release Please uses these to generate the changelog.
6. Open a pull request against `main`.

### Recovering legacy generated API docs

The docs build manages `docs/app/routes/api-reference/`. Ardo marks generated output with `.ardo-generated` and refuses to remove an existing non-empty directory without that marker. This protects handwritten files, so do not bypass the error with an automatic delete.

If an older checkout has an unmarked directory:

1. Check the working tree and make a backup outside the repository. Keep the backup until the docs build succeeds and every file has been reviewed.

    ```bash
    git status --short
    backup_dir="$(mktemp -d ../xlsx-format-api-reference-backup.XXXXXX)"
    cp -a docs/app/routes/api-reference/. "$backup_dir/"
    ```

2. Inspect the backup and separate any handwritten Markdown from generated API pages. Move handwritten pages to an authored location such as `docs/app/routes/guide/`; do not put them back inside the generated directory.
3. After the backup is verified and the directory contains no content that must be kept, move the reviewed directory aside and run the canonical docs build:

    ```bash
    mv docs/app/routes/api-reference "${backup_dir:?Run step 1 in this shell first}/api-reference-reviewed"
    pnpm --filter docs build
    ```

    The generator recreates the directory and its `.ardo-generated` marker.

Never remove an unmarked directory automatically, and never delete it before the backup and content review. A fresh checkout does not need this recovery step; `pnpm install` followed by `pnpm --filter docs build` creates the marked generated output.

## Commit Message Format

This project uses Conventional Commits to automate changelog generation:

- `feat: ...` -- new feature (minor version bump)
- `fix: ...` -- bug fix (patch version bump)
- `docs: ...` -- documentation only
- `refactor: ...` -- code change that neither fixes a bug nor adds a feature
- `test: ...` -- adding or updating tests
- `chore: ...` -- maintenance tasks

A breaking change adds `!` after the type: `feat!: remove deprecated API`

## Project Structure

```
src/
  api/        Public API functions (read, write, convert)
  ssf/        Number format engine (SpreadSheet Format)
  xlsx/       Core XLSX parsing and writing
  xml/        XML parser, writer, and escaping
  zip/        ZIP compression (CRC32, streams)
  utils/      Cell addressing, dates, buffers
  types.ts    TypeScript type definitions
  index.ts    Public exports
tests/        Vitest test files
```

## Reporting Bugs

Open an issue with a minimal reproduction. If possible, attach the `.xlsx` file that triggers the bug.
