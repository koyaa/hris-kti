# Codebase Audit Report

**Project:** KTI HRIS Quick Menu UserScript (`quick-menu.js`)  
**Date:** 2025  
**Scope:** Security, Performance, Code Organization, Technical Debt, Best Practices, Bug Risks  
**File:** `quick-menu.js` (6,673 lines)  
**Runtime:** Browser (Tampermonkey/Greasemonkey on https://hris.kti.co.id/)

---

## Executive Summary

The KTI HRIS Quick Menu is a large single-file UserScript that automates attendance management, shift assignment, employee distribution, and anomaly detection on the HRIS web application. The codebase demonstrates competent JavaScript engineering with proper async/await patterns, centralized configuration, and deliberate XSS prevention measures.

However, the monolithic architecture (6,673 lines, 188 functions) creates significant maintainability risk. The implicit cross-page state machine built on 111 sessionStorage operations is fragile and difficult to reason about. Several race conditions exist in concurrent batch processing, and hardcoded year-specific data will require annual maintenance.

### Overall Health Assessment

| Dimension | Rating | Summary |
|-----------|--------|---------|
| Security | Good | Strong XSS sanitization; minor CSRF validation gap |
| Performance | Adequate | Appropriate for typical use; batch operations could stress the server |
| Maintainability | Poor | Single 6,673-line file with high coupling |
| Code Quality | Moderate | Good patterns mixed with legacy inconsistencies |
| Reliability | Moderate | Race conditions and fragile DOM dependencies |

### Metrics at a Glance

| Metric | Count |
|--------|-------|
| Total lines | 6,673 |
| Functions | 188 (41 async) |
| sessionStorage operations | 111 |
| setTimeout calls | 55 |
| await Promise delays | 16 |
| window.location navigations | 35 |
| alert() calls | 19 |
| try/catch blocks | 28 |
| GM_getValue/setValue | 14 |
| var declarations (legacy) | 8 |

---

## 1. Security

### 1.1 XSS Prevention (Severity: Low - Well Mitigated)

The codebase employs a layered defense against cross-site scripting:

- **`parseHTML()`** (line ~340): Sanitizes HTML via DOMParser, stripping `<script>`, `<style>`, `<iframe>` tags, removing `on*` event attributes and `javascript:` URLs.
- **`setInnerHTML()`** (line ~348): Wrapper that routes all innerHTML assignments through `parseHTML()`.
- **`escHtml()`** (line ~336): HTML entity escaping for interpolated values in templates.
- **25 uses of `escHtml`** across template rendering demonstrate consistent application.
- Only **6 raw `.innerHTML` occurrences** exist, with only 2 direct assignments located inside the sanitization wrapper itself.

**Residual Risks (Severity: Low):**

- `insertAdjacentHTML('beforeend')` in `applyMark()` (around line 3235) constructs button HTML using string interpolation. Data attributes use `escHtml()`, but the element structure relies on trusted input.
- `shiftEl.innerHTML +=` in `applyDistDropdowns()` (around line 6300) appends hardcoded option elements, bypassing the `setInnerHTML()` wrapper. While currently safe (options are static strings), this bypasses the established pattern.
- The global `HTML` template literal (line ~1814, approximately 400 lines) is static trusted content but its size makes manual review difficult.

### 1.2 CSRF Handling (Severity: Medium)

**Strengths:**
- Properly extracts CSRF tokens from `<meta>` tags and hidden form fields before POST requests.
- Multiple fallback selectors for CSRF token detection demonstrate awareness of DOM variations.

**Weakness:**
- No validation that a CSRF token was actually found before submitting requests. If the token selector fails (e.g., due to HRIS page structure change), requests will proceed without CSRF protection and silently fail or succeed without authorization verification.

**Recommendation:** Add a guard clause that aborts the request and notifies the user if no CSRF token is detected.

### 1.3 Data Exposure (Severity: Low)

- No sensitive credentials (API keys, passwords) are hardcoded.
- Employee data (NRP, names, department info) is cached in `sessionStorage` with keys like `qm_jk_`, `qm_nama_`, `qm_bag_`. This data is cleared on tab close but remains accessible to any same-origin script during the session.
- `cachedEditHtml` stores full edit-page HTML in memory, which may contain form tokens.

### 1.4 Session State Security (Severity: Low)

- 111 sessionStorage operations create a complex implicit state machine across page navigations.
- No integrity validation on stored session data. Corrupted or tampered sessionStorage entries could trigger unexpected automation behavior.
- Session keys are predictable (`qm_` prefix), making them easy to identify for any same-origin code.

---

## 2. Performance

### 2.1 DOM Operations (Severity: Medium)

| Pattern | Location | Impact |
|---------|----------|--------|
| Full table rebuild per item | `renderBatchResults()` | O(n) DOM rebuild on every batch item completion (debounced 150ms) |
| Full list rebuild per check | `renderAnomalies()` | Rebuilds anomaly list on each SPKL check completion |
| Row iteration per anomaly | `applyMark()` | Iterates all table rows for each anomaly application |
| Double table scan | `scanAttendanceTable()` | Pre-pass row count + full scan = two iterations |

No virtual scrolling or pagination exists for result sets that can reach 100 NRPs. For typical usage (10-20 employees), this is acceptable. For maximum batch sizes, DOM thrashing during rapid result updates could cause UI jank.

### 2.2 Memory (Severity: Low)

- `state.batchLogs`: Accumulates up to 500 entries (hardcoded cap) with no concern.
- `cachedEditHtml`: Single HTML string in memory, manageable.
- `state.profileStats`: Accumulates indefinitely in debug mode with no upper bound.
- `state.anomalies`: Can grow without limit during extended single-page usage.

### 2.3 Network (Severity: Medium)

- **`APP_CONFIG.BATCH_POOL_SIZE = 10`** concurrent workers hit the HRIS server simultaneously. Each worker makes 2-3 sequential fetch requests (search + general + profile).
- **No request deduplication**: The same NRP can be fetched multiple times across different features (batch check, SPKL, distribution).
- Appropriate timeouts: 15-second default via AbortController; 300-second timeout for `executeBackgroundDistribusi` long-running operations.

### 2.4 Timing and Delays (Severity: Low)

- 55 `setTimeout` calls and 16 `await new Promise(r => setTimeout(r, ...))` delays exist.
- `TIMING` constants are centralized (good), but several hardcoded delays remain in callbacks: 6000ms for Bagian dropdown AJAX, 1200ms in `fillDistribusiForm`.
- `fillDistribusiForm` uses a fixed 6-second timeout waiting for a dropdown AJAX response rather than polling or event-driven detection.
- The 150ms debounce on render is appropriate.

---

## 3. Code Organization and Structure

### 3.1 Monolithic Architecture (Severity: High)

A single 6,673-line file containing 188 function declarations makes this codebase extremely difficult to maintain, review, or onboard new contributors to. While the UserScript format constrains the deployment model (single file required), the development workflow need not be constrained to a single source file.

### 3.2 Section Layout

The file is organized with numbered section headers (0-25):

| Section | Purpose |
|---------|---------|
| 0 | Configuration (frozen constants, rules, thresholds) |
| 1 | Shared Helpers (routing, URL builders, DOM utilities) |
| 2 | Employee Fetch Helpers (`getEmp`, cache read/write) |
| 3 | Debugger/Logger |
| 4 | Styles (GM_addStyle, ~500 lines of CSS) |
| 5 | State |
| 6 | HTML Template (~400 lines) |
| 7-8 | UI Helpers and Service |
| 9-10 | Anomaly Helpers, Batch Check |
| 11-14 | SPKL Highlight, NRP Check, Panel Toggle, Anomaly Detection |
| 15-21 | Feature-specific automation (SPKL, Barcode, Distribusi, etc.) |
| 22-23 | Event Handlers, Init |

This numbered organization is reasonable within the single-file constraint and demonstrates intentional structure.

### 3.3 Coupling (Severity: High)

- The `state` object is a god object accessed by virtually every function in the file.
- UI rendering functions directly modify both DOM and `state` simultaneously, making it impossible to test logic in isolation.
- No separation between data layer, business logic, and presentation.
- The sessionStorage-based state machine tightly couples page navigation flow to automation logic.

---

## 4. Duplication and Technical Debt

### 4.1 Duplicated Patterns (Severity: Medium)

**Form serialization:** `buildFormPayload()` exists as a utility, yet `processSpklBackgroundSingle()`, `saveJkMaster()`, `saveKKMaster()`, `distributeKkBackground()`, and `executeBackgroundDistribusi()` each manually serialize forms with nearly identical code.

**Employee route selection:** `getRoute()`, `getEmployeeRouteSet()`, and `buildEmployeeUrls()` all implement similar OS/internal branching logic for employee type routing.

**Batch processing loops:** `_continueSpklBatch()`, `_processManyNrpPage()`, `_processHadirManyNrpPage()`, and `checkHadirBulanResume()` all follow an identical pattern:
```
loop → setFieldValue → await delay → click button → await delay
```

**Dropdown population:** `fillDistribusiForm()` and `pilihDropdownDinamis()` both solve dynamic dropdown population but with completely different approaches.

### 4.2 Style Inconsistencies (Severity: Low)

- **8 `var` declarations** in `renderBatchResults` / `exportBatchResults` while the rest of the file uses `const`/`let`.
- Mix of callback style (`pilihDropdownDinamis`) and async/await.
- Mix of `function name() {}` declarations and `const name = () => {}` expressions.

### 4.3 Hardcoded Values (Severity: Medium)

| Value | Location | Risk |
|-------|----------|------|
| `LIBUR_NASIONAL_2026` | Configuration section | Requires annual update; no holidays for 2027+ |
| `'2026'` | HTML template year input default | Stale after year change |
| `2020-2035` | `runBatchCheck()` year range | Will need extension |
| `500` | Inline in batch log logic | Magic number, should be in config |

### 4.4 Dead Code and Comments (Severity: Low)

- Comments reference removed functions: "fetchEmployeePage -- removed", "log(msg, data) -- removed"
- Commented-out code blocks: `// openPanel()` and `// const tabFix = ...` in `checkSpklBatchResume`

---

## 5. Best Practices

### 5.1 Error Handling (Severity: Medium)

**Strengths:**
- 28 try/catch blocks provide good coverage for async operations.
- Error states are generally communicated to the user.

**Weaknesses:**
- Some catch blocks swallow errors with only `Logger.warn` (e.g., SPKL Online check), making failures invisible to the user.
- **19 `alert()` calls** for user-facing errors exist alongside the `UI.showResult()` system, creating an inconsistent error presentation.
- No global `unhandledrejection` handler. A forgotten `await` on a rejected promise will fail silently.

### 5.2 Async Patterns (Severity: Low)

- Good use of async/await throughout the codebase.
- `Promise.all` used for parallel employee page fetches.
- Worker pool pattern (`BATCH_POOL_SIZE`) for batch processing is well-implemented.
- **Gap:** No cancellation for in-flight fetch requests when a batch is aborted. `state.batchAborted` only prevents new work from starting; existing network requests complete and their results are discarded.

### 5.3 Variable Scoping (Severity: Low)

Duplicate state exists in several places:

| Module variable | State property | Notes |
|----------------|----------------|-------|
| `shortcutKey` | `state.shortcut` | Same value stored twice |
| `alwaysCollapseMenu` | `state.alwaysCollapse` | Same value stored twice |
| `isRecordingShortcut` | (none) | Should be a state property |
| `cachedEditHtml` | `state._lastEditNrp` | NRP awareness bolted on via separate state key |

### 5.4 Naming Conventions (Severity: Low)

- Indonesian and English are mixed: `tebakShiftSebenarnya`, `pilihDropdownDinamis`, `hariValid` alongside `renderBatchResults`, `detectAnomalies`.
- Consistent `qm-` prefix for CSS classes and IDs (good).
- Constants use `UPPER_CASE` (good).
- Domain abbreviations (`KLR`, `MSK`, `OTB`, `OTL`, `JK`, `KK`) require domain knowledge to understand.

---

## 6. Bug Risks

### 6.1 Race Conditions (Severity: High)

- **`state.anomalies`** is modified by multiple concurrent async operations (`detectAnomalies`, `checkSPKLOnline`, `checkBarcodeMangkir`). While JavaScript is single-threaded, interleaved microtask scheduling means mutations from one operation can be observed mid-update by another.
- **`renderAnomalies()`** reads `state.anomalies` which may be in an inconsistent state between async yield points.
- **`state.pendingChecks`** is decremented in `finally` blocks. Multiple concurrent completions with render triggers can show stale pending counts.
- **`state.batchQueue`** is shared across workers using `shift()`. This is safe in single-threaded JS, but the debounced render (150ms) can display stale batch progress.

### 6.2 State Management (Severity: High)

- The cross-page sessionStorage state machine has **no version or schema validation**. A format change across script updates will silently break in-progress automations.
- If a page crashes mid-automation, sessionStorage retains stale state that will re-trigger automation on next page load.
- `STORAGE.AUTO_FINISHED` flag can persist indefinitely if return-URL navigation fails.
- Multiple automation features (SPKL batch, Hadir batch, Distribusi) can potentially conflict if triggered simultaneously by the user.

### 6.3 Edge Cases (Severity: Medium)

| Pattern | Risk |
|---------|------|
| `parseTime()` returns null for invalid input | Callers do not always null-check the return value |
| `isOutsourceNrp()` checks `length === 8` | Does not validate all-digit constraint |
| `countLibur()` and attendance scan | Assume table has 12+ columns; will fail on different page layouts |
| `getPageContext()` NRP detection | Multiple fallback mechanisms could return wrong NRP if page has multiple NRP references |
| `LIBUR_NASIONAL_2026` | No holidays defined for 2027+; holiday detection will silently stop working |

### 6.4 DOM Dependencies (Severity: High)

The script heavily depends on specific HRIS page structure:

- Table column indices are hardcoded in `COL` configuration.
- Form field names, button text content, and modal selectors are hardcoded.
- Multiple selector fallbacks exist (e.g., `SELECTORS.SPKL_MODAL_MSK` has 4 alternatives), suggesting HRIS layout changes are frequent.
- If HRIS updates its HTML structure, features will **silently fail** rather than alerting the user.

---

## 7. Recommendations

### Priority Matrix

| # | Recommendation | Severity | Effort | Impact |
|---|---------------|----------|--------|--------|
| 1 | Add CSRF token validation guard | Medium | Low | High |
| 2 | Implement state schema versioning for sessionStorage | High | Medium | High |
| 3 | Add DOM structure validation with user notification on mismatch | High | Medium | High |
| 4 | Centralize hardcoded year data into a maintainable calendar config | Medium | Low | Medium |
| 5 | Unify error presentation (replace `alert()` with `UI.showResult()`) | Medium | Low | Medium |
| 6 | Add request cancellation via AbortController on batch abort | Medium | Medium | Medium |
| 7 | Consolidate form serialization into shared `buildFormPayload()` | Medium | Medium | Low |
| 8 | Eliminate duplicate state variables (shortcutKey, alwaysCollapseMenu) | Low | Low | Low |
| 9 | Replace `var` with `const`/`let` in remaining 8 occurrences | Low | Low | Low |
| 10 | Remove dead code comments and commented-out blocks | Low | Low | Low |

### Detailed Recommendations

#### Immediate (Low Effort, High Impact)

1. **CSRF validation guard**: Before any POST request, check that the extracted CSRF token is non-empty. If missing, abort the request and notify the user that page structure may have changed.

2. **Year/calendar externalization**: Move `LIBUR_NASIONAL_2026` to a dynamically-named structure (e.g., `LIBUR_NASIONAL[year]`) and add a startup check that warns the user if the current year has no holiday data defined.

3. **Unify error UX**: Replace the 19 `alert()` calls with `UI.showResult()` to provide consistent, non-blocking error presentation.

#### Short-term (Medium Effort, High Impact)

4. **SessionStorage schema versioning**: Add a schema version key to sessionStorage. On script load, validate stored data against expected schema. If mismatched, clear stale state and notify the user rather than attempting to resume broken automation.

5. **DOM contract validation**: At initialization, verify critical selectors resolve to actual DOM elements. If key page structure is missing, disable affected features and show a notification rather than failing silently.

6. **Batch abort cancellation**: When `state.batchAborted` is set, call `.abort()` on all active AbortControllers for in-flight requests, rather than just preventing new work.

#### Long-term (High Effort, Structural)

7. **Build pipeline introduction**: Adopt a bundler (e.g., Rollup with a UserScript plugin) to allow multi-file development while producing the required single-file output. This enables:
   - Separation of concerns (config, helpers, features, UI)
   - Unit testing of pure logic functions
   - Linting and automated style enforcement

8. **State management refactor**: Replace the god-object `state` with a minimal reactive store or at minimum, accessor functions that can validate state transitions and prevent concurrent mutation conflicts.

9. **Request deduplication layer**: Implement a fetch cache keyed by NRP + endpoint that prevents redundant network requests across features within the same page session.

---

## Appendix: Methodology

This audit was conducted through static analysis of `quick-menu.js` (commit `c684217`). Findings reference approximate line numbers based on the numbered section layout. No dynamic analysis or runtime testing was performed due to the dependency on the live HRIS application at https://hris.kti.co.id/.

### Severity Definitions

| Level | Definition |
|-------|-----------|
| **Critical** | Exploitable vulnerability or data loss risk requiring immediate fix |
| **High** | Significant reliability or maintainability risk affecting ongoing development |
| **Medium** | Correctness or quality issue that should be addressed in the near term |
| **Low** | Minor improvement or cleanup with no immediate operational risk |

---

*End of Audit Report*
