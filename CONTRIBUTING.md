# Contributing to xlsx-format

Thanks for your interest in contributing! Here's how to get started.

## Setup

```bash
git clone https://github.com/nickelow/xlsx-format.git
cd xlsx-format
npm install
```

## Development Workflow

```bash
npm run check       # TypeScript type checking
npm run lint        # ESLint
npm run format      # Prettier (auto-fix)
npm test            # Run tests once
npm run test:watch  # Run tests in watch mode
npm run build       # Build ESM + CJS bundles
```

## Making Changes

1. Fork the repository and create a branch from `master`.
2. Write your code. Follow the existing style -- Prettier and ESLint enforce most of it.
3. Add or update tests for any changed behavior.
4. Make sure all checks pass: `npm run check && npm run lint && npm test`
5. Use [Conventional Commits](https://www.conventionalcommits.org/) for your commit messages (e.g. `feat: add X`, `fix: handle Y`). Release Please uses these to generate the changelog.
6. Open a pull request against `master`.

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
