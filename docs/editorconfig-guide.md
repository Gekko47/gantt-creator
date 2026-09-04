# EditorConfig and .gitattributes guide

> Reference for the cross-tool configuration files at the repository root.
> See also `docs/work-items/R0.4-editorconfig-and-gitattributes.md`.

## `.editorconfig`

The `.editorconfig` file configures code style and analyzer severity
for the Gantt Creator solution. It applies to:

- Visual Studio 2022 17.14+ and Visual Studio 2026
- VS Code with the C# Dev Kit extension
- JetBrains Rider 2024.3+
- `dotnet format` in CI (R0.7)

The root has `root = true`, so settings here are not inherited from
any parent `.editorconfig`.

### Sections

| Section | What it controls |
| --- | --- |
| `[*]` | Universal settings: UTF-8, LF, space indent, trim trailing whitespace, final newline |
| `[*.md]` | Preserves trailing whitespace (Markdown two-space line breaks) |
| `[*.ps1]` | PowerShell: 4-space indent |
| `[*.json]` | JSON: 2-space indent |
| `[*.{yml,yaml}]` | YAML: 2-space indent |
| `[*.cs]` | C# language conventions and analyzer severities |
| `[*.Generated.cs]` | Marks generated code so analyzers skip it |
| `[tests/**/*.cs]` | Test-file-specific analyzer suppressions |

### C# language conventions

- `csharp_style_namespace_declarations = file_scoped:warning` —
  all new code uses file-scoped namespaces. Existing `namespace { }`
  blocks are left alone unless touched.
- `csharp_style_var_for_built_in_types = true:suggestion` —
  `var` is preferred for `int`, `string`, `bool`, etc.
- `csharp_style_var_elsewhere = false:suggestion` — explicit type
  is preferred when the type is not built-in.
- `dotnet_style_qualification_for_field = false:suggestion` —
  no `this.` prefix on fields and properties.
- `csharp_style_expression_bodied_properties = true:suggestion` —
  one-line properties use expression bodies.

### Analyzers

- `category-Style.severity = warning` (and Naming, Performance,
  Reliability, Security, Usage, Design) — every analyzer in these
  categories is treated as a build warning, which `TreatWarningsAsErrors`
  promotes to a build error.
- `[tests/**/*.cs]` suppresses the rules the xUnit template's
  `UnitTest1.cs` triggers. R0.5 replaces the placeholders; after
  that, the suppressions can be removed in a follow-up commit.

### Naming rules

- Private fields use `_camelCase` (the `dotnet_naming_rule.private_fields_underscore`
  block). This is the .NET runtime team's convention and matches
  the BCL style.

## `.gitattributes`

The `.gitattributes` file declares per-file-type attributes for Git.
The most important rule is line-ending normalization.

### Line endings

- `*.cs`, `*.csproj`, `*.props`, `*.targets`, `*.sln`, `*.slnx`,
  `*.editorconfig`, `*.gitattributes`, `*.gitignore` — `text eol=lf`
- `*.ps1`, `*.psm1`, `*.psd1`, `*.json`, `*.yml`, `*.yaml`, `*.xml`,
  `*.config`, `*.lock.json`, `*.ruleset` — `text eol=lf`
- `*.md` — `text eol=lf` with `whitespace=-trailing-space`
  (preserves Markdown line breaks)

### Merge strategy

- `packages.lock.json` (and the per-project copies under
  `**/packages.lock.json`) use `merge=union`. This keeps both
  versions on conflict, which is the correct behaviour for a
  deterministic pin: a conflict means a deliberate change on
  both branches and must be resolved manually.

### Binary file policy (documented for R10.2)

The `.gitattributes` documents the expected policy for binary
assets that will be added when the Excel-DNA packaging lands:

- `*.xll`, `*.dll`, `*.pdb`, `*.exe` — `binary`
- `*.png`, `*.jpg`, `*.pdf`, `*.zip` — `binary`

These rules are commented out until the assets exist. When the
Excel-DNA build produces the first `.xll`, uncomment the relevant
lines in a single commit.

## `Directory.Build.props` analyzer policy

The `NoWarn` list in `Directory.Build.props` is the authoritative
suppression list. The `.editorconfig` `[tests/**/*.cs]` section is
a belt-and-braces companion, but MSBuild's `NoWarn` is what actually
suppresses diagnostics at build time.

The `NoWarn` list is reviewed in R0.5 once the `UnitTest1.cs`
placeholders are replaced. Each rule that is removed from the list
must have a recorded reason (a fixed placeholder, a documented
exception, or a renamed method).
