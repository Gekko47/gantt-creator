# ADR-0071 — SkiaSharp as the raster renderer's drawing dependency

- **Status:** Accepted
- **Date:** 2026-11-23
- **Work item:** `R8.2` ([`docs/work-items/R8.2-skia-primitive-renderer.md`](../work-items/R8.2-skia-primitive-renderer.md))
- **Related:** ADR-0004 (Excel-DNA pins), ADR-0069 (typed PowerPoint PIA), R10.2 (release packaging), R10.7 (SBOM freeze)

## Context

Phase 8 adds the clipboard-independent 300-DPI PNG channel. The raster renderer must draw the
same immutable scene the live and editable renderers consume, on a CPU surface with pinned
antialiasing so the golden-image policy can compare exact pixels
([`docs/02-ARCHITECTURE.md`](../02-ARCHITECTURE.md) "PNG renderer"). The dependency materially
shapes output fidelity, so the dependency policy
([`docs/05-GIT-QUALITY.md`](../05-GIT-QUALITY.md)) requires an ADR before the pin lands.

R8.2's Step 0 probed the candidates on 2026-11-23 against nuget.org metadata and the Excel host:

- **SkiaSharp 4.157.0 + SkiaSharp.NativeAssets.Win32 4.157.0** — MIT; managed `lib/net10.0`
  asset; native `libSkiaSharp.dll` under `runtimes/win-{x64,x86,arm64}/native`.
- `System.Drawing.Common` 11.0.0 — MIT, but GDI+ antialiasing and DPI behaviour depend on host
  settings, which the golden-image policy cannot pin.
- `SixLabors.ImageSharp` 4.1.3 + `Drawing` 3.1.3 — pure managed, but the file-based Six Labors
  split licence is a commercial-use risk and Drawing lags the core version.
- Direct2D/WIC — raw COM interop, the highest-risk layer this repository already polices.
- WPF `RenderTargetBitmap` — needs a dispatcher inside the Excel STA add-in.

## Decision

### D1 — SkiaSharp 4.157.0, pinned centrally

`SkiaSharp` and `SkiaSharp.NativeAssets.Win32` are pinned at **4.157.0** in
[`Directory.Packages.props`](../../Directory.Packages.props) and referenced by
`GanttCreator.Raster` only. SkiaSharp is a **package** reference, not a project reference, so
Raster's Core-only project graph is unchanged and the architecture suite's dependency-direction
facts keep passing.

### D2 — CPU raster surface with pinned antialiasing

The renderer uses an `SKBitmap` surface (`SKColorType.Bgra8888`, `SKAlphaType.Premul`) and pins
`SKPaint.IsAntialias = true`; no GPU surface, clock, random, or culture input participates in a
render. `RasterDeterminism` records the surface kind, the antialias setting, and the SkiaSharp
assembly version beside the pixels.

### D3 — `SKPathBuilder`, not the obsolete `SKPath` mutators

In 4.157.0 `SKPath.MoveTo`/`LineTo`/`Close` are obsolete; the probed replacement is
`SKPathBuilder` (`AddPoly`/`Detach`). Lines use `SKCanvas.DrawLine` directly. No renderer API is
written from memory; each is verified against the installed package's XML documentation.

### D4 — The native binary ships beside the packed XLL

Excel-DNA packs managed assemblies into the XLL, but SkiaSharp's default loader does not resolve
the packed `NATIVE_LIBRARY_LZMA` resource: an in-host `AutoOpen` render threw
`TypeInitializationException` on `SkiaApi`. With `libSkiaSharp.dll` placed **beside** the packed
XLL the same in-host render passed (`OK:4.157.0.0`). The chosen deployment route is
**beside-the-XLL**: R10.2 copies the native binary into the release payload and R10.7 records the
MIT licence in the SBOM/third-party-notice freeze. The R8.2 stop condition ("native assets cannot
ship in the add-in payload") is therefore not triggered.

## Consequences

- The PNG channel gets a deterministic CPU renderer that can draw hatches and dashes more
  faithfully than the host in R8.3.
- Two release obligations are now explicit and owned: R10.2 must copy `libSkiaSharp.dll`, and
  R10.7 must record the MIT notice. Dropping either breaks the shipped add-in.
- Golden baselines are reference-machine artefacts; the R8.2 byte-identity pin proves
  same-machine determinism, and cross-machine equality is neither claimed nor tested until R8.4's
  representative baselines.
- A future SkiaSharp bump is a dependency-policy change with its own probe and an amendment to
  this ADR.

## Alternatives considered

The Step-0 table above records why each candidate lost. In summary: SkiaSharp is the only option
that is MIT, ships a `net10.0` asset, rasterizes on the CPU with pinnable antialiasing, encodes
PNG, and has a proven in-host load route.
