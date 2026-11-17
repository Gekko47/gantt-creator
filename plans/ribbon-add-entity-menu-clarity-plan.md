# Add-entity split buttons — menu clarity plan

**Status:** implemented 2026-11-17. Owner approved the layout, the 17 keytips and the 17-row wording;
the Office host pass is owed.
**Work item:** the **surface revision of [`R5.16a`](../docs/work-items/R5.16a-ribbon-layout-and-icons.md)**,
which already owns the Data group's layout, `menuSeparator` and the icon map. No new roadmap row.
**Nature:** reversible presentation. No schema change, no state-model change, no new dependency.

---

## 1. What the three menus are today (read from source)

| Fact | Evidence |
|---|---|
| Three `splitButton`s in the Data group: `spbAddActivity`, `spbAddMilestone`, `spbAddElements` | [`Ribbon.xml`](../src/GanttCreator.AddIn/RibbonResources/Ribbon.xml:31) |
| Item counts are 9 / 5 / 3, ids are `ent` + the `GanttEntityType` name | [`Ribbon.xml`](../src/GanttCreator.AddIn/RibbonResources/Ribbon.xml:33) |
| Only the activity menu has a divider, and it is the only titled one (`title="Procurement"`) | [`Ribbon.xml`](../src/GanttCreator.AddIn/RibbonResources/Ribbon.xml:40) |
| Item labels are projected from `EntityTypeCatalog`, never authored in XML | [`GanttRibbon.cs`](../src/GanttCreator.AddIn/GanttRibbon.cs:446) |
| Each primary shows the last menu pick, label *and* glyph | [`RibbonStateService.cs`](../src/GanttCreator.AddIn/RibbonStateService.cs:2445) |
| Menus are static XML — a `menu` is not bindable in Excel-DNA 1.9 | [`R5.3-wi-e-add-child.md`](../docs/work-items/R5.3-wi-e-add-child.md:121) |
| 9 activity rows share one glyph (`add-activity`); 5 milestone rows share `add-milestone` | [`RibbonGlyphCatalog.g.cs`](../src/GanttCreator.AddIn/RibbonGlyphCatalog.g.cs:72) |
| Screentips restate the label; supertips restate the label again | [`Ribbon.xml`](../src/GanttCreator.AddIn/RibbonResources/Ribbon.xml:34) |
| Keytip schemes are per-menu inventions: `AP…AS`, `BP…BU`, `JD/JS/JP`; primaries `AA/AM/AD` | [`Ribbon.xml`](../src/GanttCreator.AddIn/RibbonResources/Ribbon.xml:47) |
| Tests already pin the family partition (both directions), the projected primary label, keytip uniqueness (1–3 chars) and every supertip | [`AddRowRibbonTests.cs`](../tests/GanttCreator.AddIn.Tests/AddRowRibbonTests.cs:204), [`GanttRibbonTests.cs`](../tests/GanttCreator.AddIn.Tests/GanttRibbonTests.cs:939) |

## 2. Review findings

| # | Problem | Why it matters in Excel | Severity |
|---|---|---|---|
| P0 | The split-button primary **renames itself to the last menu pick**, so the menu's **first block has no stable heading** and the top of the list is never described | The owner's report: after picking Spacer then reopening, the button says Spacer and the top block is unexplained; the button cannot be the menu's label (owner ruling 2026-11-17) | High |
| P1 | Two of the three menus have **no grouping at all**: 5 milestones and 3 elements are one undivided list; the activity menu's single heading covers only the last 3 of 9 rows | A 9-row dropdown with no headings is a scanning task, not a choice; the one heading that exists proves the mechanism works | High |
| P2 | **No visual separator between concepts** in `mnuAddElements`: a chart marker (Delineator, reads `Start`, draws a full-height line) sits beside two table-structure rows (Spacer, Splitter, no plot geometry) | Users pick Spacer expecting a line, or Delineator expecting a blank row | High |
| P3 | **Icons do not discriminate rows**: all 9 activity rows carry the same 16 px glyph, all 5 milestone rows the same | At 16 px the picture adds nothing, so the label carries 100% of the decision | Medium |
| P4 | **Supertips repeat the label** ("Adds an as-planned activity row.") instead of stating the differentiator (dates read, style requirement, parent rule, chart vs table) | The one place with room to explain *when* to choose an item says nothing new | Medium |
| P5 | **Keytip families are inconsistent**: milestones use a `B` prefix that also means Baseline, elements use `J`, primaries `AA/AM/AD` mix two schemes | Keytips cannot be learned; `AD` is the element primary but `JD` is its first item | Low |
| P6 | Ordering is **already consistent** with the entity guide (Planned → Built → Baseline → Critical → Custom) and with the split-button defaults | Not a defect — recorded so no reordering churn is introduced | — |

## 3. Proposed menu layout (Slice 1)

This is the owner's layout of 2026-11-17, adopted as written. The keytips keep the owner's scheme with five codes corrected in §3.1. Ordering is unchanged; only dividers and keytips change, so the state model, ids and callbacks are untouched.

```text
Add activity (AA)                        Add milestone (AM)                Add elements (AD)
------------------------------           ------------------------          ------------------------------
-- Activity --                           -- Milestone --                   -- Chart marker --
As-Planned Activity        APA           As-Planned Milestone     APM     Delineator             DEL
As-Built Activity          ABA           As-Built Milestone       ABM     ------------------------------
Baseline Activity          BLA           Baseline Milestone       BLM     -- Table structure --
Custom Activity            CUA           Critical Milestone       CRM     Spacer                 SPA
------------------------------           ------------------------          Splitter               SPL
-- Events --                             -- Custom --
Delay Event                DLY           Custom Milestone         CUM
Critical Interval          CRI
------------------------------
-- Procurement --
As-Planned Procurement     APP
As-Built Procurement       ABP
Baseline Procurement       BLP
```

Rules the layout follows:

1. **A heading names a concept, never a type.** Headings: `Activity`, `Events`, `Procurement`, `Milestone`, `Custom`, `Chart marker`, `Table structure`.
2. **Every section carries a heading, including the first** (owner ruling 2026-11-17). The primary renames itself to the last menu pick (R5.18g), so the button cannot say what the top block is; an unlabelled first block is exactly the ambiguity reported. A leading `menuSeparator` with a `title` is the only heading a `menu` offers.
3. **Keytips encode basis + kind in three characters** (owner ruling 2026-11-17, reviewed in §3.1). The first two letters are the basis (`AP`, `AB`, `BL`, `CU`, `CR`) and the last letter is the kind (`A` activity, `P` procurement, `M` milestone, `I` interval). Three is the `keytip` attribute's maximum, so the scheme replaces the current two-letter mixture (`AP`, `BP`, `JD`) rather than extending it. All 17 were checked against the 41 keytips already on the tab: **no exact, case-insensitive duplicate**, and the primaries keep `AA`, `AM`, `AD` (no new code begins with those pairs, so no primary gains a longer prefix rival).

### 3.1 Keytip review — owner values to proposed values

The owner's scheme is right where it follows basis + kind. Five codes break its own rule, so they are corrected here; the other twelve stand.

| Row | Owner | What the code actually says | Proposed | Why |
|---|---|---|---|---|
| As-Planned / As-Built / Baseline / Custom Activity | `APA` `ABA` `BLA` `CUA` | `AP` `AB` `BL` `CU` + `A` activity | unchanged | Reads exactly as intended; the four share a kind letter, which is correct |
| As-Planned / As-Built / Baseline Procurement | `APP` `ABP` `BLP` | basis + `P` procurement | unchanged | Correct, and the `P` ending is what separates them from the activity rows |
| As-Planned / As-Built / Baseline Milestone | `APM` `ABM` `BLM` | basis + `M` milestone | unchanged | Correct |
| Custom Milestone | `CUM` | `CU` + `M` | unchanged | Correct |
| Critical Interval | `CRI` | `CR` + `I` interval | unchanged | Correct |
| Critical Milestone | `CIM` | `C` + `I` + `M`, i.e. it spends a letter on **Interval** | `CRM` | `CRM` is `CR` + `M`, the exact counterpart of `CRI`; `CIM` reads as a third, different concept |
| Delay Event | `DLE` | `D` + `L` + `E`, no basis/kind structure, and it starts with the existing `DL` keytip | `DLY` | `DLY` reads as "delay" and keeps three characters. The `DL` overlap is harmless only because menu keytips are scoped to the open menu (see below) |
| Delineator | `TD` | an invented `T` prefix with no link to the word | `DEL` | `DEL` is the display name's own abbreviation and matches the rest of the scheme |
| Spacer | `TS` | invented `T` prefix | `SPA` | As above |
| Splitter | `TP` | invented `T` prefix | `SPL` | As above; `SPA`/`SPL` differ at the third character, which is the one the user types last |

Two behavioural notes that the live pass must confirm:

- **Menu-item keytips are scoped to the open menu** (inference, not a repo fact). If that holds, `DLY` sharing `DL` with the Labels group's Description-label dropdown cannot slow that control down, and `DEL`/`SPA`/`SPL` cannot interfere with the ribbon-level `D`, `C`, `S`-family keytips. If the live check disproves it, the fallback is to keep the owner's two-letter keys for the element menu rather than introduce a prefix rivalry.
- **No new code is a prefix of another within the same menu.** Within Activity, `APA`/`APP` differ at the third character; within Milestone, `APM`/`ABM`/`BLM`/`CRM`/`CUM` differ by the second; within Elements, `SPA`/`SPL` differ at the third. Excel shows the candidates and the last keystroke decides, which is the designed behaviour rather than an ambiguity.
4. **One heading per real conceptual boundary**, and only where the boundary matters: planned/built/baseline/custom are one decision axis; delay and critical interval are another; procurement is a different visual (hatched) entirely.
5. **The style-key rows are called out.** `Custom Activity` and `Custom Milestone` require a `StyleKey` (chosen in the Labels group), so the milestone menu gets a `Custom` heading rather than burying the caveat in a tooltip.
6. **Element menu splits chart from table.** Delineator draws at a date; Spacer and Splitter are table structure with no plot geometry.

## 4. Supertip rewrite (Slice 2)

Today every entity row's screentip restates the label ("Add an As-Planned Activity") and its supertip restates it again ("Adds an as-planned activity row."). Both are wasted: the row's label is visible beside the tooltip, so the only useful job for the tooltip is to answer **how do I choose between these**.

Five rules for the rewrite:

1. **The screentip names the concept, not the action.** "Planned span", not "Add an As-Planned Activity".
2. **The supertip carries at most two facts**, in this order: what the row draws, then the one thing that changes the choice (dates read, style-key requirement, parent rule, chart versus table).
3. **Only guide facts.** Every claim traces to a §, listed in the last column so the wording can be re-checked when the guide changes. Nothing about behaviour the guide does not state.
4. **No sentence starts "Adds …".** The button already says Add; the tooltip says what the thing is.
5. **No links.** Supertips are local text only, per the offline invariant.

Length: the supertip stays under about 120 characters, because the host wraps a longer one onto a third line and pushes the tip off the row's visual context.

| Row | Screentip | Supertip | Source |
|---|---|---|---|
| As-Planned Activity | Planned span | The planned programme span. Reads Start and Finish; may share its lane with actual and baseline rows. | §13 |
| As-Built Activity | Actual span | The actual execution span. Reads Start and Finish; critical intervals point at it by parent. | §14, §16 |
| Baseline Activity | Baseline span | The approved baseline span. Reads Start and Finish; the stack index decides which bar sits where. | §15 |
| Custom Activity | Style-defined span | Needs a style key from the Labels group. Reads Start and Finish; an unknown key is an error, never a fallback. | §19 |
| Delay Event | Delay span | An explicit delay period. Reads Start and Finish; drawn in its own red body with the label inside by default. | §17 |
| Critical Interval | Critical portion | Reads its own Start and Finish. Attaches to the nearest activity that can own it; disabled when none can. | §16 |
| As-Planned Procurement | Planned procurement | A planned procurement period. Reads Start and Finish; drawn hatched instead of solid. | §18 |
| As-Built Procurement | Actual procurement | An actual procurement period. Reads Start and Finish; drawn hatched instead of solid. | §18 |
| Baseline Procurement | Baseline procurement | A baseline procurement period. Reads Start and Finish; drawn hatched instead of solid. | §18 |
| As-Planned Milestone | Planned milestone | A planned milestone. A four-point diamond at its Start date; Finish is never read. | §20, §21 |
| As-Built Milestone | Actual milestone | An actual milestone. A four-point diamond at its Start date; Finish is never read. | §20, §21 |
| Baseline Milestone | Baseline milestone | A baseline milestone. A four-point diamond at its Start date; Finish is never read. | §20, §21 |
| Critical Milestone | Critical milestone | A critical milestone. A four-point diamond at its Start date; it takes no styling from a nearby interval. | §20, §21 |
| Custom Milestone | Style-defined milestone | Needs a style key from the Labels group. A diamond at its Start date; an unknown key is an error, never a fallback. | §19, §21 |
| Delineator | Full-height date line | A full-height line at its Start date. Finish, lane and stack are kept but flagged as not used. | §24 |
| Spacer | Blank spacer row | A blank row that only adds vertical space. No chart shape; any dates on it are ignored with a warning. | §11 |
| Splitter | Section band | A named band across the table and plot that separates sections. No date geometry; the description is the band label. | §10 |

Two wording decisions worth stating, because they look like omissions:

- **"Finish is never read"** appears on every milestone rather than "uses Start only". The guide is emphatic that a milestone and a delineator read `Start` and never `Finish` for geometry, and the failure mode this prevents is a user typing a finish date and wondering why the diamond did not move.
- **No row promises an automatic result.** The tooltips do not say the chart updates, because the coalesced refresh is described once in the Refresh button's own tooltip and repeating it on 17 rows spends the space the differentiator needs.

## 5. Optional slices

| Slice | Change | Cost / risk |
|---|---|---|
| S3 keytips (**committed**) | The 17 three-letter keytips in §3, replacing `AP/AB/AL/AU/AV/AI/AQ/AR/AS`, `BP/BB/BL/BI/BU`, `JD/JS/JP` | Low. `R5.16` and `R5.16a` docs name some old values and need the mapping; the global-uniqueness test already exists and must be re-run |
| S4 glyphs (**still optional**) | Add only the concept glyphs that survive 16 px, as reviewed in §5.1 | Low if limited to two slugs. New slugs, palette regeneration, `check-ribbon-glyphs.ps1`, and a live 16 px legibility pass |

S4 is independent of S1–S3 and can be deferred without weakening them.

### 5.1 Glyph review — what S4 could actually draw

The constraint is not taste, it is raster. Every menu row rasterises to **16 px**, and the 24-unit viewbox means a 1.5-unit stroke lands on roughly one pixel. A detail that reads in a mockup disappears at that size, so the test for each candidate is whether the **silhouette** changes, not whether the drawing is prettier.

| Row(s) | Proposed slug | Geometry | Survives 16 px? |
|---|---|---|---|
| As-Planned, As-Built, Baseline Activity | `add-activity` (existing) | three offset bars, already in [`gen-ribbon-glyphs.ps1`](../scripts/gen-ribbon-glyphs.ps1:87) | Yes — keep |
| All five milestones | `add-milestone` (existing) | flag on a pole | Yes — keep |
| Delineator, Spacer, Splitter | existing three slugs | line with arrows; band | Yes — keep |
| The three Procurement rows | `procurement` (new) | the bar motif drawn as **two parallel strokes** (outline instead of solid), which is how the chart shows hatch | Yes — outline versus solid is a silhouette change |
| Critical Interval, Critical Milestone | `critical` (new, or `critical-span` plus `critical-milestone`) | the bar or flag with **one wide diagonal stroke** across it | Borderline — only if the slash is 2 units; to be judged on the live pass |
| Delay Event | `span-delay` (new) | the bar motif with a **gap** in one bar | No — the gap closes to sub-pixel; rejected |
| Custom Activity, Custom Milestone | `style-defined` (new) | a small pen or star accent on the bar or flag | No — the accent is noise at 16 px; rejected |

### 5.2 Alternatives for Delay Event and Custom that survive 16 px

The rule for the two rows the owner asked about: at 16 px on a 24-unit viewbox a glyph survives only if the **silhouette** changes — stroke count, orientation, a closed shape, or a gap of three units or more. Decoration (dots, small accents, hatch) never survives.

| Row | Candidate | Geometry | Why it survives | Verdict |
|---|---|---|---|---|
| Delay Event | `span-delay` | one bar split in two with the right half dropped about four units (`M3 10h8`, `M12 15h9`) | bar count drops from three to one and the halves are visibly stepped; four units is about 2.7 px at 16 px | **Recommended** |
| Delay Event | `hourglass` | Lucide `hourglass` | closed two-triangle shape, unique on the tab; reads as elapsed time | Alternate — accurate about delay, but it does not depict a span under an activity |
| Delay Event | hairline gap in one bar | `M3 12h7`, `M11 12h10` | gap is about 1.3 px and closes in raster | Rejected |
| Custom Activity and Custom Milestone | `pen` | Lucide `pen`, one diagonal nib path | the only diagonal-dominant glyph on the tab, so nothing competes with it; "you define the style" is what the pen says | **Recommended** |
| Custom Activity and Custom Milestone | `circle-dashed` | Lucide `circle-dashed` | an open ring reads as "not defined yet"; the dashes may merge but the ring still reads | Alternate if the pen reads as a generic slash |
| Custom Activity and Custom Milestone | `palette` | Lucide `palette` | the blob survives, but its main path is already the Theme settings glyph, so it duplicates a silhouette already on the tab | Rejected |
| Custom Activity and Custom Milestone | `swatch-book` | stacked rounded rectangles | reads as `show-columns` or `initialise-sheet` at 16 px | Rejected |
| Custom Activity and Custom Milestone | accent dot or star on the bar or flag | — | sub-pixel at 16 px | Rejected |

Sharing one glyph between Custom Activity and Custom Milestone is deliberate: they are the same concept (a style key must be chosen) in two menus, and each sits under a `Custom` heading that already names it. The existing design already shares one glyph across the family.

**Owner ruling 2026-11-17: adopt the recommendation.** Fund `procurement`, `span-delay` and `pen` (three slugs, six PNGs), and add `critical` only if the live pass says its diagonal survives. Do not chase per-row uniqueness for planned versus built versus baseline: in the chart those three are the same bar with a different fill, so giving them different 16 px silhouettes would misdescribe the product, and nine near-identical bars with tiny differences is worse for scanning than nine identical ones. R5.19's own gate already exempts menu rows from slug uniqueness, which is that decision already recorded.

**Cheaper alternative worth a ruling:** drop glyphs from the entity menu rows entirely and let the headings and labels carry the menu. Excel menus routinely ship text-only rows, and this removes P3 instead of decorating it. It does reverse R5.19's deliberate choice, so it is not proposed as the default.

## 6. Acceptance tests

**Automated (Slice 1 + 2):**

- Menu-shape contract test: for each of `mnuAddActivity`, `mnuAddMilestone`, `mnuAddElements`, assert the exact child sequence of item ids and separators, that the **first child of every menu is a titled separator**, that every separator carries a non-empty `title` from the approved vocabulary, and that every separator carries no `keytip` or `onAction`. A separator does require an `id` (the schema demands one), so the assertion is on the two channels that would make it behave like a control.
- Every declared section title is in the approved vocabulary (`Events`, `Procurement`, `Custom`, `Table structure`) — no free text.
- Existing pins stay green and unmodified: family partition both ways, primary `getLabel` (no static `label`), keytip uniqueness and the 1–3 character limit, glyph coverage, supertip presence.
- A keytip pin that states the 17 expected values by id, so the scheme cannot drift back to a per-menu mixture.
- `check-md-links.ps1`, `check-status.ps1`, pre-commit hook.

**Live Office (human, `verify-office.ps1` plus a manual menu pass):**

- Each of the three menus opens with its headings rendered as group headers, no clipping, no horizontal scroll, and the expected row count (9 / 5 / 3 items plus separators).
- The split-button primary still adds the item it names, and the label/glyph follow a menu pick.
- `Critical Interval` still greys out on a table with no eligible parent.

Slices 1–2 touch XML and tests only; no C# and no workbook schema, so no scene/refresh gate applies.

## 7. Out of scope (and why)

- **Rewording item labels away from catalogue display names** — that would create a second text authority beside `EntityTypeCatalog`, and display names are schema-bound. Headings supply the missing context instead.
- **Dynamic menu items** — not possible in Excel-DNA 1.9 (recorded in `R5.3` WI-E).
- **A tick on the current primary's item** — the item is an add action, not a toggle; a pressed state on it reads as "on/off". The column-visibility menu earned its ticks because those rows *are* toggles.
- **Putting a verb in the primary label** ("Add As-Planned Activity") — lengthens a large button and duplicates the menu rows' intent.
- **Reordering to match the Type dropdown order** — the dropdown is schema order; the menus are task order. Recorded, not changed.
- **The Fix split button's one-item menu (`Validate`)** — an owner-tooled decision from R5.16, not part of this review.

## 8. Drift found while reviewing (needs an owner ruling, not a silent fix)

`EntityTypeCatalog` holds **17** selectable types ([`EntityTypeCatalog.cs`](../src/GanttCreator.Core/EntityTypeCatalog.cs:28)), including `Custom Milestone`, but:

- the entity guide's Type table lists **16** rows and omits `Custom Milestone` ([`07-GANTT-ENTITY-GUIDE.md`](../docs/07-GANTT-ENTITY-GUIDE.md:139));
- the unknown-type validation message says "use one of the **16** catalogue display names" ([`GanttRowValidator.cs`](../src/GanttCreator.Core/GanttRowValidator.cs:257)).

The ribbon menus already offer all 17. This conflict sits between source-of-truth items 2, 3 and 7, so it stops for a decision: correct the guide row and the message in the same work item (separate `docs:` commit) or record it as a known drift.

## 9. Suggested landing order

1. `docs:` work item `R5.21` + this plan, with the section-title vocabulary recorded as an owner ruling.
2. `feat(ribbon):` Slice 1 dividers + the menu-shape contract test.
3. `docs(ribbon):` Slice 2 supertips + the supertip wording pin.
4. Optional `feat(ribbon):` S3 keytips, then S4 glyphs (own commit each, S4 needs the live 16 px pass).
5. `docs:` catalogue-count drift, if the owner approves the correction.
