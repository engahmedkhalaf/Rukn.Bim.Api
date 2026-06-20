# QicBoqMapper — Bug Fix & Refactor Implementation Plan

_Date: 2026-06-16 · Scope: `QIC 5D BOQ Code Generator/QicBoqMapper`_

This plan is grounded in a full read of the source. Items are ordered by
severity. Each lists the symptom, root cause (with file:line), and the fix.

---

## Phase 1 — Correctness bugs (do first)

### 1.1 "Matching Method" dropdown does nothing
- **Symptom:** Selecting `Element ID` vs `Category + Family + Type` has no effect.
- **Cause:** `PerformMapping` ([QicBoqManager.cs:193](QicBoqMapper/QicBoqManager.cs))
  and `UpdateParameters` ([QicBoqManager.cs:324](QicBoqMapper/QicBoqManager.cs))
  take `matchingMethod` but never read it. Both always try Element ID first,
  then fall back to Category+Family+Type.
- **Fix (decided: strict, mutually exclusive modes):** Branch on `matchingMethod`.
  - `"Element ID"` → match **only** by the id lookup; no type fallback.
  - `"Category + Family + Type"` → match **only** by the composite key; **no**
    Element ID fallback.
  - Apply the same branch in both `PerformMapping` and `UpdateParameters` so they
    stay consistent (or, better, have `UpdateParameters` reuse the match already
    decided by `PerformMapping` — see 3.1). Build only the dictionary the chosen
    mode needs.

### 1.2 In-window license activation is a disconnected mock — remove it
- **Symptom:** The main window's *Activate License* shows "License: Active",
  yet **Validate** and **Generate** stay disabled.
- **Cause:** `ActivateLicense` ([QicBoqGeneratorViewModel.cs:604](QicBoqMapper/QicBoqGeneratorViewModel.cs))
  only checks `LicenseKey.Length >= 5` and sets a label; it never calls
  `LicenseManager.SaveLicense`. But `CanValidateOrGenerate`
  ([QicBoqGeneratorViewModel.cs:466](QicBoqMapper/QicBoqGeneratorViewModel.cs))
  gates on `LicenseManager.IsActivated()`, which reads the registry.
- **Fix (decided: remove the mock):** The License **ribbon button** (real
  Supabase flow via `LicenseWindow` → `LicenseManager`) becomes the single
  activation path. Steps:
  1. Delete `ActivateLicense`, `ActivateLicenseCommand`, `LicenseKey`, and the
     mock `LicenseStatus` setter usage from the ViewModel.
  2. Remove the corresponding activation controls (key textbox + button) and
     their bindings from `QicBoqGeneratorWindow.xaml`.
  3. Keep a **read-only** status indicator initialized from
     `LicenseManager.IsActivated()` (refresh it when the window is activated /
     on Idling) so the label reflects real registry state.
  4. Verify the build has no dangling XAML bindings after removal.

### 1.3 `SetParameterValue` can abort the whole Generate transaction
- **Symptom:** Generate fails with an exception for some models; nothing is
  written (transaction rolls back).
- **Cause:** `param.Set(value)` ([QicBoqManager.cs:429](QicBoqMapper/QicBoqManager.cs))
  is called whenever `!param.IsReadOnly`, with no storage-type check. If the
  loose name-variant lookup (`GetParameterNameVariants`,
  [QicBoqManager.cs:433](QicBoqMapper/QicBoqManager.cs)) lands on a non-text
  parameter, `Set(string)` throws and aborts the transaction for all elements.
- **Fix:** Guard with `param.StorageType == StorageType.String` before setting,
  and wrap each set in try/catch so one bad element can't roll back the batch.

### 1.4 Loose parameter-name matching can write to the wrong parameter
- **Symptom:** Values occasionally land on an unintended (or type-level)
  parameter shared by all instances of a type.
- **Cause:** `GetParameterNameVariants` generates many spellings
  (spaces, title case, camelCase, lowercase) and `SetParameterValue` also
  searches Type parameters ([QicBoqManager.cs:410-425](QicBoqMapper/QicBoqManager.cs)).
- **Fix:** Match against the exact canonical names created in
  `CreateSharedParameters` (instance bindings). Drop the speculative variants,
  or restrict the variant set to the documented column aliases only.

### 1.5 License dialog message contradicts the rule
- **Cause:** `ValidateInput` requires code length `>= 4`
  ([LicenseManager.cs:132](QicBoqMapper/LicenseManager.cs)) but the error says
  "min 8 characters" ([LicenseWindow.xaml.cs:28](QicBoqMapper/LicenseWindow.xaml.cs)).
- **Fix:** Make the message match the rule (pick one minimum, use it both places).

---

## Phase 2 — Robustness & reporting

### 2.1 Centralize separator handling
- **Cause:** ViewModel stores the *display* string (`"- (Dash)"`, `". (Dot)"`)
  and passes it straight to `PerformMapping`, which re-parses it with fragile
  `Contains(".")`/`Contains("-")` logic ([QicBoqManager.cs:288-293](QicBoqMapper/QicBoqManager.cs)).
  Persistence is also inconsistent (default `"Dash"` vs list `"- (Dash)"`,
  remapped in the ctor at [QicBoqGeneratorViewModel.cs:67](QicBoqMapper/QicBoqGeneratorViewModel.cs)).
- **Fix:** Add one helper (e.g. `SeparatorService.ToChar(style)`) used by both
  the VM and manager; persist a single canonical token.

### 2.2 Clarify Generate summary counts
- **Cause:** Summary shows `Elements Updated: {MatchedTypesCount}` /
  `Elements Skipped: {ErrorCount}` ([QicBoqGeneratorViewModel.cs:553](QicBoqMapper/QicBoqGeneratorViewModel.cs)).
  `MatchedTypesCount` includes **Warning** rows (matched but missing fields),
  which *are* written with partial data — so the two-line summary hides warnings.
- **Fix:** Report Success / Warning / Error separately; rename the misleading
  `*TypesCount` properties (they count elements, not types).

### 2.3 Guard `ToDictionary` against duplicate keys
- **Cause:** `auditRecords.ToDictionary(a => a.ElementId)`
  ([QicBoqManager.cs:326](QicBoqMapper/QicBoqManager.cs)) throws on any dup.
- **Fix:** Use a grouping/`ToDictionary` with last-wins, or assert uniqueness.

### 2.4 EPPlus licensing — keep EPPlus, license it correctly (decided)
- **Cause:** `ExcelPackage.LicenseContext = NonCommercial` is set
  ([QicBoqExcelService.cs:12](QicBoqMapper/QicBoqExcelService.cs),
  [QicBoqExportService.cs:13](QicBoqMapper/QicBoqExportService.cs), and
  `QicBoqAuditExporter.cs`), but this is a **paid, licensed** add-in (Supabase
  activation), so the NonCommercial context is not valid for distribution.
- **Decision: stay on EPPlus — no library rewrite.** Resolve the license only.
  1. Acquire a commercial EPPlus license (Polyform/commercial key from the EPPlus
     vendor) and record it in the project for reference.
  2. Replace the `LicenseContext = NonCommercial` calls with the commercial
     license registration appropriate to the EPPlus version in use — verify the
     installed version first ([QicBoqMapper.csproj](QicBoqMapper/QicBoqMapper.csproj)):
     - EPPlus 5–7: `ExcelPackage.LicenseContext = LicenseContext.Commercial;`
     - EPPlus 8+: `ExcelPackage.License.SetCommercial("<key>");`
  3. Centralize this in one place (a small static initializer) instead of
     repeating it in each service, so the key/context is set once.
- **Out of scope (per decision):** no migration to ClosedXML or any other Excel
  library; reader/exporters keep their current EPPlus code.

---

## Phase 3 — Refactor (after behavior is correct & verified)

### 3.1 Single matching pass
- `PerformMapping` and `UpdateParameters` each rebuild `idDict`/`typeDict` and
  re-derive type/family per element ([QicBoqManager.cs:202-211 & 330-339](QicBoqMapper/QicBoqManager.cs)).
  Have `PerformMapping` attach the matched `BoqRecord` to each `AuditRecord`
  (add a field), so `UpdateParameters` consumes it directly — no second lookup.

### 3.2 Replace magic strings with constants/enums
- Status `"Success"/"Warning"/"Error"`, selection modes
  (`"Current Selection"/"Active View"/"Entire Model"` with a stray
  `"Whole Model"` alias at [QicBoqGeneratorViewModel.cs:79](QicBoqMapper/QicBoqGeneratorViewModel.cs)),
  and matching methods should be enums/`const`.

### 3.3 Remove dead/unused plumbing
- `matchingMethod` is threaded everywhere but unused (resolved by 1.1); audit
  for other unused params once matching is fixed.

---

## Suggested execution order
1. Phase 1 (1.1 → 1.5) — small, isolated, high impact.
2. Build + smoke test in Revit: load a sample BOQ xlsx, run Validate, run
   Generate on a small selection, export the audit report.
3. Phase 2, re-test.
4. Phase 3 refactor, relying on the now-verified behavior.

## Validation checklist
- [ ] `Element ID` mode matches **only** by id; `Category + Family + Type` mode
      matches **only** by type — neither falls back to the other.
- [ ] In-window activation controls removed; License ribbon button is the only
      activation path and a read-only status reflects `IsActivated()` on open.
- [ ] Generate completes without rollback on a model containing unrelated
      same-named parameters of non-text type.
- [ ] Separator output matches the dropdown for all four styles.
- [ ] Generate summary numbers reconcile (Success + Warning + Error = processed).
- [ ] EPPlus runs under a commercial license (no `NonCommercial` context), set
      once centrally; BOQ load + element export + Excel audit export still work.
