# Contributing to OfficeForge

Thanks for helping out. This guide covers setup, the checks every change must pass, and how to open a pull request.

## Prerequisites

- **.NET SDK 10.0.x** (the projects target `net10.0`). Check with `dotnet --version`.
- Git.

## Clone

```bash
git clone https://github.com/Sarmkadan/officeforge.git
cd officeforge
```

## Build

```bash
dotnet build                 # Debug, whole solution (OfficeForge.slnx)
dotnet build -c Release      # the configuration CI uses
```

## Run tests

```bash
dotnet test                                        # all xUnit tests
dotnet test --filter "FullyQualifiedName~CellRefTests"   # a single test class
```

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) before making structural changes.

## Code style

- There is no `.editorconfig` in the repo. Match the surrounding code, and run `dotnet format` on the files you touch.
- `Nullable` and `ImplicitUsings` are enabled in every project. Fix nullable warnings in code you change; do not add new ones.
- Warnings should be treated as errors. `TreatWarningsAsErrors` is not set yet, so the build will not fail on them. Do not add new warnings.
- Public API needs XML doc comments. Guard arguments with `ArgumentException.ThrowIfNullOrWhiteSpace` and similar helpers.
- Library models must not reference OpenXML types. Use invariant culture for number and date formatting and parsing.
- Tests are named `MethodOrScenario_Condition_Expected` and use Arrange/Act/Assert comments.

## Branch naming

Create a branch from the default branch (`master`). Use a short, lowercase, hyphenated name with a type prefix that matches the commit type:

```
feat/<short-description>      e.g. feat/docx-table-export
fix/<short-description>       e.g. fix/xlsx-shared-strings-null
docs/<short-description>
chore/<short-description>
```

## Commit messages

Use [Conventional Commits](https://www.conventionalcommits.org/) prefixes:

- `feat:` new behaviour
- `fix(scope):` bug fix, with an optional scope such as `fix(xlsx):`
- `docs:` documentation only
- `chore:` tooling, repo maintenance, or dependency updates

Write the subject in the imperative mood, lowercase, and under about 72 characters. Add a body when the reason for the change is not obvious from the diff.

## Pull request checklist

Before you open or update a PR, confirm each item:

- [ ] `dotnet build -c Release` succeeds with no new warnings.
- [ ] `dotnet test -c Release` passes.
- [ ] New public API has XML doc comments, and new behaviour has tests.
- [ ] Commit messages follow the conventions above.
- [ ] The PR description explains **what** changed and **why**, and lists anything a reviewer should look at closely (format limitations, breaking changes, follow-up work).
- [ ] The PR is focused on one change. Unrelated formatting or refactoring goes in a separate PR.
