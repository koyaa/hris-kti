# Learnings & Conventions
## minimize-menu-draggable

### Codebase Conventions
- Vanilla HTML/CSS/JS — no frameworks, no build tools
- CSS custom properties in `:root` for all design tokens
- `--transition-fast: 180ms ease` for all transitions
- `.is-hidden` utility class uses `display: none !important`
- Existing JS pattern: DOM queries at top, functions below, event listeners after
- Icons stored as SVG markup strings in `iconMarkup` object

### Design Tokens
- Shadows: `--shadow-panel`, `--shadow-card`
- Radius: `--radius-panel: 22px`, `--radius-card: 12px`, `--radius-pill: 10px`
- Spacing: `--space-2` through `--space-12`
- Panel max width: `--panel-max-width: 820px`

### File Map
- `index.html`: HTML structure (82 lines)
- `app.js`: All JS logic (295 lines) — data arrays, iconMarkup, DOM queries, render functions, event handlers
- `styles.css`: All styles (566 lines) — :root tokens, layout, components, responsive breakpoints
22#NW|- Persisted minimize state belongs between the data arrays and `iconMarkup`; `safeStorage` should wrap `localStorage` so private-browsing failures never bubble up

### CSS Added (Task 1)
- `.floating-menu-btn` and `.minimize-pill` rules appended at line 568+ (after responsive breakpoints)
- Focus-visible outline uses `rgba(122, 165, 194, 0.42)` matching existing codebase pattern (not a CSS var)
- `.minimize-pill` box-shadow uses `0 1px 0 rgba(255, 255, 255, 0.85) inset` matching `.footer-button` pattern
- `.floating-menu-btn:hover` uses `translateY(-1px) scale(1.04)` for subtle lift+grow effect
- `.minimize-pill:hover` uses `translateY(-1px)` matching `.footer-button:hover` pattern

### Task 2 Learnings
- Minimize pill button inserted as static HTML in `.footer-actions` after `.nav-button-group` and before `.footer-divider`
- `iconMarkup.chevronDown` added after `book` key (alphabetical position) — two-path V shape, stroke-width 1.8
- LSP servers (biome, typescript-language-server) not installed in this environment — verified JS syntax via `node -c` and `new Function()` instead
- No click handlers added (Task 4 scope) — button is purely structural at this stage

### Task 4 Learnings
- Minimize/restore toggle should only swap `.is-hidden` on `.command-menu` and manage the floating button DOM; it must not clear search input or results state.
- The floating restore button can be recreated on demand with `iconMarkup.gridDots`, and its position should come from persisted `menuState.buttonEdge` / `menuState.buttonTop` values.

### Task 5 Learnings
- Floating button drag is wired from `createFloatingButton()` immediately after `document.body.appendChild(button)`, with module-level `dragState` before the function.
- Pointer drag clamps vertical movement with `window.innerHeight - 48`, saves `menuState.buttonTop`, and persists `{ edge, top }` through `saveState()` after snap.
- A capture-phase click listener calls `preventDefault()` and `stopImmediatePropagation()` when `dragState.moved` is true so the existing toggle click handler still handles ordinary clicks but not post-drag synthetic clicks.
- Browser verification was blocked because Chrome/Chromium is not installed; behavior was verified with a Node DOM harness covering right snap, left snap, clamp, save, drag-click suppression, and non-drag restore.

### Task 6 Learnings
- `setupDrag(button)` belongs directly in `createFloatingButton()` after `document.body.appendChild(button)` so the draggable handlers are attached immediately on creation.
- Task 6 integration fix: hide `.command-menu` before the initial render when minimized, clamp the floating button on `resize`, and add `tabIndex` plus Enter/Space handling for the restore button.

## Scope Fidelity Check - 2026-05-12

VERDICT: APPROVE. Reviewed plan Must NOT Have plus full app.js, styles.css, and index.html. No out-of-scope changes found: quickActions/resultTabs/results data unchanged, existing icon entries preserved with only chevronDown and gridDots added, existing render/filter/escapeHtml functions unchanged, CSS minimize rules appended after existing stylesheet, and index.html only adds one .minimize-pill button inside .footer-actions. Repo surface contains no package/build tooling or external libraries. LSP diagnostics were unavailable because configured servers are not installed; node --check app.js passed.


### Code Quality Review - 2026-05-12
- Verdict: REJECT
- Verification: `node --check app.js` passed with no syntax errors.
- Blocking issues:
  - `app.js:315-345` only clears drag state and `document.body.style.userSelect` on `pointerup`; missing `pointercancel`/lost-capture cleanup can leave dragging state/body selection disabled after canceled pointer interactions.
  - `styles.css:603`, `styles.css:625`, and `styles.css:638` use hardcoded `rgba(...)` values in the new minimize section, conflicting with the review requirement that CSS colors come through `var(--*)` tokens.


## Manual QA - minimize-menu-draggable - 2026-05-12

VERDICT: REJECT

Environment: Playwright with user-local Chromium from `/tmp/opencode`; system Chrome was unavailable, and Chromium required user-space library extraction under `/tmp/opencode/browser-root`. Tested `file:///home/ree/proyek/icikiw/index.html` with real browser clicks, typing, keyboard, pointer drag, reload, and resize.

Scenario results:
- S1 PASS: `.minimize-pill` rendered with `aria-label="Minimize menu"`, one chevron SVG, and `Minimize` text. Click added `.is-hidden` to `.command-menu`; `.floating-menu-btn` appeared fixed on the right edge (`right=16px`, rect left/right approx `936/984` at 1000px viewport).
- S2 PASS: Clicking `.floating-menu-btn` restored `.command-menu` and removed the floating button from the DOM.
- S3 PASS: Search value `farm` and Farmlink results persisted after minimize/restore.
- S4 PASS: Console path `menuState.minimized = true; saveState();` wrote `localStorage.getItem('menu-minimized') === "true"`; reload opened hidden menu with floating button and `menuState.minimized === true`.
- S5 FAIL: During pointer drag, vertical movement is clamped, but horizontal movement does not move the button with the pointer and snap edge does not update when dragged past the midpoint. Evidence: dragging from right edge toward `x=10` left the button on the right (`floatingLeft=936px`, `floatingRight=16px`, rect left/right approx `935/985`) and `menuState.buttonEdge` stayed `right`.
- S6 PASS: Plain click on minimized floating button restored menu instead of drag-snapping.
- S7 PASS: Resize clamps after the CSS `top` transition completes. Recheck evidence: before resize viewport `1000x800`, styleTop `720px`; after resizing to height `360` and waiting 350ms, styleTop `312px`, rect top/bottom `312/360`, `menuState.buttonTop=312`.
- S8 PASS: Tab focus reached `.floating-menu-btn` with visible outline (`solid 2px`); Enter restored menu. Re-minimize + Tab + Space also restored menu.

Failing repro for S5:
1. Open `index.html`.
2. Run `menuState.minimized = true; showFloatingButton();` in console.
3. Press and hold the floating button on the right edge.
4. Drag downward, then drag horizontally past the left/right midpoint toward the left edge, and release.
5. Expected: button follows pointer horizontally, snaps to the left edge, and `menuState.buttonEdge` becomes `left`.
6. Actual: button only changes vertical `top`; it remains on the right edge and `menuState.buttonEdge` remains `right`.
### Task 7 Learnings
- Horizontal dragging now needs `startX`/`startLeft` alongside the existing Y fields, and `pointercancel` should mirror `pointerup` cleanup by restoring `document.body.style.userSelect`.

## Manual QA Re-run S5/S6 - 2026-05-12

VERDICT: APPROVE

Environment: Playwright Chromium from `/tmp/opencode` with local dependency path `/tmp/opencode/browser-root/usr/lib/x86_64-linux-gnu`, viewport `800x600`, opened `file:///home/ree/proyek/icikiw/index.html`. Test setup used `menuState.minimized = true; menuState.buttonEdge = 'right'; menuState.buttonTop = 200; showFloatingButton();`.

Scenario results:
- S5 PASS left snap: slow diagonal drag down-left past midpoint released with `menuState.buttonEdge === "left"`; button rect `left=16`, `right=64`, `top=336`, `bottom=384`, inline `left: 16px`, `top: 336px`.
- S5 PASS top/left clamp: drag beyond top-left clamped inline position to `left: 0px`, `top: 0px` during drag.
- S5 PASS right snap: drag back to the right released with `menuState.buttonEdge === "right"`; button rect approximately `left=735.04`, `right=784.96`, inline `right: 16px`, `top: 476px`.
- S5 PASS bottom/right clamp: drag beyond bottom-right clamped inline position to `left: 752px`, `top: 552px` during drag and remained on the right edge after release.
- S6 PASS click without drag: movement under 5px restored the menu; after click `menuState.minimized === false`, `.floating-menu-btn` was removed, `.command-menu` was visible, and `buttonEdge` remained `right` (no snap).
