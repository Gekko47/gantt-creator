# ADR-0070 — Pinned-font provisioning for raster text and golden images

- **Status:** Accepted
- **Date:** 2026-11-23
- **Work item:** `R8.2a` ([`docs/work-items/R8.2a-font-pinning-adr.md`](../work-items/R8.2a-font-pinning-adr.md))
- **Related:** [`ADR-0071`](0071-skia-raster-renderer-dependency.md) (SkiaSharp), R8.3 (text/hatch rendering), R8.4 (representative goldens), R10.7 (SBOM/third-party notices), [`docs/02-ARCHITECTURE.md`](../02-ARCHITECTURE.md) ("Coordinate and rounding policy"), [`docs/04-TEST-STRATEGY.md`](../04-TEST-STRATEGY.md) ("Golden images"), [`docs/07-GANTT-ENTITY-GUIDE.md`](../07-GANTT-ENTITY-GUIDE.md) (typography tokens and §22), `AGENTS.md` (proprietary-fonts stop condition)

## Context

Phase 8 adds the clipboard-independent 300-DPI PNG channel. The raster renderer
consumes the same immutable scene as the live and editable channels, and R8.3
draws the scene's text. The scene already measures every label with the
injected [`AptosTextMetrics`](../../src/GanttCreator.Core/Scene/AptosTextMetrics.cs)
over the committed [`FontAdvanceTable`](../../src/GanttCreator.Core/Scene/FontAdvanceTable.cs),
which [`scripts/gen-font-advances.py`](../../scripts/gen-font-advances.py)
generated from the **real font binaries**. The four files it read are:

- `Aptos/30153066857.ttf` (Regular)
- `Aptos/31531277363.ttf` (Bold)
- `Aptos Narrow/27970306811.ttf` (Regular)
- `Aptos Narrow/24328775610.ttf` (Bold)

The advance table's shared denominator is `UnitsPerEm = 2048`. The raster
renderer must therefore draw the **same four faces** the measurement table was
generated from; a different face can make the measured box and the drawn glyphs
disagree in the direction that clips a label. The architecture pins the
text-metrics service's "implementation, font, and version" for tests
([`02-ARCHITECTURE.md`](../02-ARCHITECTURE.md), "Coordinate and rounding
policy"), and the golden-image policy pins font family and font version
([`04-TEST-STRATEGY.md`](../04-TEST-STRATEGY.md), "Golden images").

The entity guide sets the typography token to `Aptos` and requires that a
missing font produce "a clear error or approved fallback, never an unreported
substitution" ([`07-GANTT-ENTITY-GUIDE.md`](../07-GANTT-ENTITY-GUIDE.md),
typography tokens and §22). The manifest decision `D-G4` recorded this choice
as licence-sensitive and reserved to the human
([`IMPLEMENTATION-PLAN.md`](../work-items/IMPLEMENTATION-PLAN.md), decision
register).

### The observed font-route constraint (reference machine, 2026-11-23)

A Step-0 probe against the pinned SkiaSharp 4.157.0 package established the
capability boundary that decides the route:

```text
SKFontManager.Default.GetFontFamilies()    -> 208 families; Aptos absent
SKFontManager.Default.MatchFamily("Aptos") -> null
SKTypeface.FromFamilyName("Aptos")         -> Segoe UI (silent fallback)
SKTypeface.FromFile(<cloud-font path>)     -> Aptos / Aptos / Version 2.01;O365
```

Office cloud fonts live under
`%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts\<Family>` and are **not**
registered with the OS font system (zero Aptos entries in `HKCU` and `HKLM`).
SkiaSharp's system font manager therefore cannot see them, and the convenience
lookup silently substitutes Segoe UI — exactly the outcome the entity guide
forbids. The only faithful SkiaSharp route is the explicit typeface loader
(`SKTypeface.FromFile` / `FromStream`; `SKFontManager.CreateTypeface(Stream)`)
plus an identity check.

## Decision

### D1 — Option B: use the Office-installed Aptos faces in place; do not bundle

The raster text path resolves the four faces from the Office cloud-font cache
through the explicit SkiaSharp typeface loader. No font file is copied into the
repository, the add-in payload, or any build output.

**Licence evidence (primary sources).** Microsoft Learn's Aptos page lists
redistribution under a separate commercial licence and points to the Font
redistribution FAQ and "License Microsoft fonts … for hardware & software
redistribution or server installations"
([Aptos font family](https://learn.microsoft.com/typography/font-list/aptos)).
Microsoft's font-licensing guidance states that most licences for the fonts
Microsoft provides with applications and Windows "don't let you put the fonts in
applications or share them with others"
([Font Licensing for the brand center](https://learn.microsoft.com/sharepoint/brand-font-licensing)).
The Font redistribution FAQ limits its guidance to fonts supplied with Windows
and directs readers to the licence agreements that accompany other applications
for fonts supplied with them
([Font redistribution FAQ](https://learn.microsoft.com/typography/fonts/font-faq)).
Office cloud fonts are documented as a Microsoft 365 feature
([Cloud fonts in Office](https://support.microsoft.com/office/cloud-fonts-in-office-f7b009fe-037f-45ed-a556-b5fe6ede6adb)).

**Conclusion.** No redistribution right for Aptos was established from these
sources. Bundling (Option A) is therefore rejected under the `AGENTS.md`
proprietary-fonts stop condition; redistributing it would require a separate
commercial licence. Option B does not redistribute the font: it reads the copy
Office installed for the user's own Office installation and keeps the font
bytes out of the repository and the release payload.

Option C (an approved open fallback) is **not** adopted. No open font is known
to be Aptos-metric-compatible, and the committed `FontAdvanceTable` is
Aptos-specific; a fallback would make the PNG's text differ from both Excel and
the measured label boxes. It remains a last resort only if Option B cannot be
made deterministic.

### D2 — Pin four per-face identities, not one family version

The pin is the four identities below. The versions are per face: the family's
Regular and Bold faces differ (2.01 vs 2.00), so a single family version would
be a false pin.

| Face | File | Family | PostScript name | `name` version | `head` revision | upem | SHA-256 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Aptos Regular | `30153066857.ttf` | `Aptos` | `Aptos` | `Version 2.01;O365` | 2.0100 | 2048 | `95980114FCFD42F2F9C446DAE429B70582BF2F03097D68433EA9E7D85A49DA0B` |
| Aptos Bold | `31531277363.ttf` | `Aptos` | `Aptos-Bold` | `Version 2.00;O365` | 2.0000 | 2048 | `B4124A6C17E73189574D9E51ADDA221F236D1DF901F1BBCD57E9A09AE5B18599` |
| Aptos Narrow Regular | `27970306811.ttf` | `Aptos Narrow` | `Aptos-Narrow` | `Version 2.01;O365` | 2.0100 | 2048 | `DED7515AA064578485A09C11A798A05D0303BDACDD7F2DE92D0521F3BC7793BA` |
| Aptos Narrow Bold | `24328775610.ttf` | `Aptos Narrow` | `AptosNarrow-2bc0000-560000` | `Version 2.00;O365` | 2.0000 | 2048 | `D1E33B7FFA6C481063DF87516D0C2976DCE00C717663C7943C2742E3E4BCAA99` |

`UnitsPerEm = 2048` is the cross-check against the measurement authority: a
loaded face with a different denomination cannot be the face `FontAdvanceTable`
was generated from. The `name` table's ID 5 and `head.fontRevision` are read
through `SKTypeface.TryGetTableData`; SkiaSharp exposes no version property.

### D3 — Fail closed; never the system font manager in the text path

R8.3 resolves each face through `PinnedFontSource` (the single resolver) and
validates the identity **before the first text draw**. `SKFontManager.MatchFamily`
and `SKTypeface.FromFamilyName` are forbidden in the text path: the probe proves
they return null or Segoe UI for the pinned family. A missing or mismatched face
is a typed refusal (`FontFileMissing`, `LoadFailed`, `FamilyMismatch`,
`PostScriptMismatch`, `VersionMismatch`, `UnitsPerEmMismatch`, `ParseFailure`);
the render aborts and no partial PNG with substituted text is produced.

### D4 — Golden-image policy: pinned identity, reference-machine baselines, human approval

Every golden baseline records the four identity records above beside the pinned
renderer, OS, culture, DPI, and antialias settings. The harness asserts the live
run's identity equals the recorded baseline identity. Baselines are
reference-machine artefacts: a baseline update is a separate commit, human
approved, whose message states why every changed image is correct; baselines are
never regenerated automatically in CI
([`04-TEST-STRATEGY.md`](../04-TEST-STRATEGY.md), "Golden images").

A missing pinned face on a non-reference host has one of two honest outcomes and
never a silent substitution:

- **Production renderer:** the typed refusal in D3; no PNG is emitted.
- **Font-dependent test:** the test is skipped with a stated reason and the
  environment recorded. GitHub-hosted `windows-latest` has no Office cloud
  fonts, so the routine CI gate cannot execute font-dependent goldens; the
  reference-machine run is the gate. This is recorded as a known limitation,
  not a green coverage claim.

The R8.2a probe is the executable form of this policy. It lives at
[`tests/GanttCreator.Raster.Tests/FontProbeTests.cs`](../../tests/GanttCreator.Raster.Tests/FontProbeTests.cs)
and every test carries `[Trait("Category", "FontProbe")]`. The routine entry
points exclude that category
([`scripts/test-non-office.ps1`](../../scripts/test-non-office.ps1),
[`scripts/verify.ps1`](../../scripts/verify.ps1)), pinned by
[`scripts/ci-parity.Tests.ps1`](../../scripts/ci-parity.Tests.ps1). The
reference-machine evidence command is:

```powershell
dotnet test tests/GanttCreator.Raster.Tests/GanttCreator.Raster.Tests.csproj `
  -c Release --filter "Category=FontProbe" --logger "console;verbosity=detailed"
```

### D5 — One pin authority; a font change is a measured change

`PinnedFontSet` is the single pin and `PinnedFontSource` the single resolver
(R8.3). The probe and the renderer share both, so a probe that disagrees with
the renderer is structurally impossible. Because the four faces are exactly the
files `gen-font-advances.py` read, changing a face or version requires
regenerating `FontAdvanceTable` and a reviewed golden update — never a silent
font swap.

## Probe evidence (2026-11-23)

Observed on the reference machine with SkiaSharp 4.157.0. The probe passes five
facts; the record and the negative control are:

```text
Aptos/30153066857.ttf        family='Aptos'        ps='Aptos'                       name='Version 2.01;O365' head=2.0100 upem=2048 sha256=95980114FCFD42F2F9C446DAE429B70582BF2F03097D68433EA9E7D85A49DA0B
Aptos/31531277363.ttf        family='Aptos'        ps='Aptos-Bold'                  name='Version 2.00;O365' head=2.0000 upem=2048 sha256=B4124A6C17E73189574D9E51ADDA221F236D1DF901F1BBCD57E9A09AE5B18599
Aptos Narrow/27970306811.ttf family='Aptos Narrow' ps='Aptos-Narrow'                name='Version 2.01;O365' head=2.0100 upem=2048 sha256=DED7515AA064578485A09C11A798A05D0303BDACDD7F2DE92D0521F3BC7793BA
Aptos Narrow/24328775610.ttf family='Aptos Narrow' ps='AptosNarrow-2bc0000-560000'  name='Version 2.00;O365' head=2.0000 upem=2048 sha256=D1E33B7FFA6C481063DF87516D0C2976DCE00C717663C7943C2742E3E4BCAA99

system font families=208; contains Aptos=False
negative control: FromFamilyName("Aptos") -> Segoe UI/SegoeUI; refused as FamilyMismatch
```

The five facts are: the four-face identity pin; the system-lookup negative
control; the missing-face typed refusal; the version-mismatch positive; and the
shared `UnitsPerEm` cross-check. The run recorded `Total tests: 5, Passed: 5`.

## Consequences

- R8.3 implements `PinnedFontSet`/`PinnedFontSource`, text drawing, and the
  extension of `RasterDeterminism` with the resolved identity. It consumes this
  ADR; it does not revisit the option.
- R8.4 records the identity in each baseline's metadata. Cross-machine pixel
  equality is not claimed beyond the reference machine; the identity pin is
  what makes a different machine's failure legible.
- R10.2/R10.7 are unchanged by this decision: no font file ships and no new
  dependency is added. The SkiaSharp MIT notice remains the only new
  third-party obligation (ADR-0071).
- An Office font update changes a face's version/hash. The probe fails loudly
  and the pin, the advance table, and the affected baselines are updated
  together, with human review. That is the intended early-warning behaviour.
- No redistribution occurs and no font is embedded in the add-in; the
  proprietary-fonts stop condition is respected.

## Alternatives considered

- **Option A — bundle a font.** Rejected: the primary-source licence research
  establishes no redistribution right for Aptos, and the `AGENTS.md`
  proprietary-fonts stop condition forbids shipping it without one. A future
  purchased redistribution licence could reopen this option with a new ADR.
- **Option C — approved open fallback.** Rejected for fidelity: no open font is
  known to be Aptos-metric-compatible, and the advance table is Aptos-specific.
  A fallback would change the PNG's text relative to Excel and the measured
  boxes. Kept as a last resort only.
- **SkiaSharp's system font manager (`MatchFamily`/`FromFamilyName`).**
  Rejected on evidence: it cannot see Office cloud fonts and silently returns
  Segoe UI, which is the forbidden substitution.
- **Excel COM / VSTO / Open XML SDK / `System.Drawing`.** Rejected: none yields
  the exact face bytes **and** version; COM/VSTO add a host dependency and
  `System.Drawing` adds GDI+ determinism and dependency regressions.

## Approval

Approved 2026-11-23 with the implementation plan
([`plans/R8.2a-font-pinning-adr-plan.md`](../../plans/R8.2a-font-pinning-adr-plan.md)),
which records the four owner rulings this ADR depends on: the explicit
typeface-loader route, Option B, the probe location, and the reference-machine
golden/test policy. `D-G4` is closed in the manifest register.
