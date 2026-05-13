# Minimize Menu to Draggable Floating Button

## TL;DR
> **Summary**: Add a pill-shaped minimize toggle to the command menu that collapses it into a single draggable floating button (grid icon). The button snaps to left/right screen edges, persists position and state via localStorage, and restores the full menu on click.
> **Deliverables**: Minimize pill toggle in footer, floating draggable button with snap-to-edges, minimize/restore toggle behavior, localStorage persistence, viewport resize handling, click-vs-drag threshold, keyboard accessibility
> **Effort**: Short
> **Parallel**: YES - 2 waves
> **Critical Path**: Task 1 → Task 2,3 → Task 4,5 → Task 6

## Context
### Original Request
> "Minimize the menu to small buttons that can be moved by the user"
### Interview Summary
- **Minimize behavior**: Collapse full menu panel into a single floating round button (like chat widget collapse). Click to re-expand.
- **Minimize trigger**: Pill-shaped toggle placed in the menu footer area
- **Floating button icon**: Grid/dots icon
- **Drag behavior**: Snap to left/right screen edges; free vertical movement
- **State persistence**: localStorage remembers minimized/expanded state and button position
- **Test strategy**: Agent QA via Playwright browser verification only (no test framework added)
- **Mobile/responsive**: Must work on touch devices
### Metis Review (gaps addressed)
- Drag-vs-click threshold: prevent accidental restore after dragging (mitigation: <5px movement threshold)
- Viewport resize handling: clamp restored position to viewport bounds on load and resize (mitigation: clamp on init and resize event)
- Touch support: use Pointer Events instead of mouse events (mitigation: pointerdown/pointermove/pointerup)
- Keyboard accessibility: floating button must be focusable, Enter/Space to restore (mitigation: native `<button>` with aria-label)
- localStorage failures: wrap in try/catch, default to expanded/right-edge state (mitigation: safeStorage helpers)
- Search/result state preservation: minimize hides panel only, does not reset input or results (mitigation: toggle visibility only)

## Work Objectives
### Core Objective
Add a minimize/restore toggle to the command menu that collapses it into a single draggable floating button with edge-snap behavior and persisted state.

### Deliverables
1. Pill-shaped minimize toggle button rendered in the menu footer
2. Floating round button with grid icon, positioned at right screen edge initially
3. Drag behavior: vertical free movement, horizontal snap to left or right edge
4. Click-to-restore: clicking floating button re-expands full menu
5. localStorage persistence for minimized state and button position
6. Edge case handling: drag-vs-click threshold, viewport resize clamping, touch support, keyboard a11y

### Definition of Done
- `localStorage` contains `menu-minimized` and `menu-button-pos` keys when state is saved
- Clearing localStorage removes the keys; running `localStorage.getItem('menu-minimized')` in console returns expected value
- Floating button appears when pill toggle clicked; menu panel hidden
- Clicking floating button restores full menu; floating button hidden
- Dragging floating button vertically moves it; releasing near left edge snaps left, otherwise snaps right
- On browser resize, floating button stays within viewport bounds
- `localStorage.getItem('menu-minimized') === 'true'` after minimize; `=== 'false'` after restore

### Must Have
- Minimize pill toggle in footer that hides menu and shows floating button
- Floating button: round, ~48px, grid icon, fixed positioning, z-index above content
- Drag: pointer events, snap-to-edges, drag-vs-click threshold (5px)
- Restore: click floating button → hide it, show menu
- localStorage: persist minimized state and button position (edge + top)
- Keyboard: floating button is focusable, Enter/Space restores
- Viewport clamping on resize

### Must NOT Have
- NO external libraries or dependencies (pure vanilla JS/CSS)
- NO changes to existing quick actions, search, results logic
- NO removal of existing menu functionality
- NO auto-minimize based on screen size or inactivity
- NO animation library cruft — use CSS transitions only (var(--transition-fast))
- NO new HTML files or build tooling

## Verification Strategy
> ZERO HUMAN INTERVENTION - all verification is agent-executed.
- Test decision: none (no test framework) — QA via Playwright browser automation
- QA policy: Every task has agent-executed Playwright scenarios
- Evidence: .sisyphus/evidence/task-{N}-{slug}.{ext}

## Execution Strategy
### Parallel Execution Waves
> Target: 5-8 tasks per wave. <3 per wave (except final) = under-splitting.

Wave 1: CSS styles (Task 1), JS state management (Task 2), HTML pill toggle (Task 3) — all independent
Wave 2: Minimize/restore logic (Task 4), Drag behavior (Task 5) — depend on Wave 1
Wave 3: Integration + edge cases (Task 6) — depends on Wave 2

### Dependency Matrix
| Task | Blocks | Blocked By |
|------|--------|------------|
| 1. CSS styles | 4,5 | — |
| 2. JS state module | 4,5,6 | — |
| 3. Pill toggle | 4,5 | — |
| 4. Minimize/restore | 6 | 1,2,3 |
| 5. Drag behavior | 6 | 1,2,3 |
| 6. Integration + edge cases | — | 4,5 |

### Agent Dispatch Summary
- Wave 1: 3 tasks (visual-engineering, quick, visual-engineering)
- Wave 2: 2 tasks (quick, unspecified-high)
- Wave 3: 1 task (unspecified-high)
- Final: 4 review agents

## TODOs
> Implementation + Test = ONE task. Never separate.
> EVERY task MUST have: Agent Profile + Parallelization + QA Scenarios.

- [x] 1. Add CSS styles for floating button and minimize pill toggle

  **What to do**: Add all CSS for the floating button (`.floating-menu-btn`) and the minimize pill toggle (`.minimize-pill`) to `styles.css`. The floating button is a fixed-position round button (48px, circular, grid icon, shadow, z-index 9999) positioned at the right edge initially. The pill toggle is an inline element in the footer with a chevron-down icon and label text. Both use existing design tokens (colors, radius, shadows, transitions). Include `pointer-events: auto`, `touch-action: none` on the floating button for drag.

  **Must NOT do**: Do not add JS logic. Do not modify existing selectors. Do not change any existing layout. Do not use hardcoded pixel values for positioning in CSS — position is set via JS inline styles.

  **Recommended Agent Profile**:
  - Category: `visual-engineering` - Reason: Pure CSS with existing design tokens, visual polish needed
  - Skills: [`frontend-ui-ux`] - needed for design token usage and polished visual output

  **Parallelization**: Can Parallel: YES | Wave 1 | Blocks: [4,5] | Blocked By: []

  **References**:
  - Pattern: `styles.css:242-257` - `.quick-action` card styling with shadow and transitions
  - Pattern: `styles.css:430-464` - `.footer-button` pattern for pill toggle sizing
  - Token: `styles.css:1-37` - all CSS custom properties to reuse
  - Token: `styles.css:36` - `--transition-fast: 180ms ease` for all transitions

  **Acceptance Criteria**:
  - [ ] `.floating-menu-btn` rule exists with: border-radius: 50%, width 48px, height 48px, position fixed, z-index 9999, pointer-events auto, touch-action none
  - [ ] `.floating-menu-btn` uses var(--shadow-panel) for box-shadow
  - [ ] `.minimize-pill` rule exists with display inline-flex, align-items center, gap, padding, border-radius matching footer button style
  - [ ] All new CSS uses variables from :root (no hardcoded colors)
  - [ ] No existing selectors modified

  **QA Scenarios**:
  ```
  Scenario: CSS compiles without errors
    Tool: Bash
    Steps: Read styles.css and verify braces are balanced, no obvious syntax errors
    Expected: File is valid CSS, new rules present
    Evidence: .sisyphus/evidence/task-1-css-styles.txt
  ```

  **Commit**: YES | Message: `style(menu): add floating button and minimize pill CSS` | Files: [styles.css]

- [x] 2. Add JS state management module for minimize feature

  **What to do**: Add a state management object (`menuState`) at the top of `app.js` after data arrays, before DOM queries. Include: `minimized` (boolean), `buttonEdge` ('left'|'right'), `buttonTop` (number). Add `safeStorage` helper with `get(key, fallback)` and `set(key, value)` methods wrapping localStorage in try/catch. Add `saveState()` and `loadState()` functions. Default state: `{ minimized: false, buttonEdge: 'right', buttonTop: 200 }`. Keys: `menu-minimized`, `menu-button-pos`.

  **Must NOT do**: Do not add DOM manipulation, event listeners, or modify any existing function. Place before `const searchInput = ...` (line ~182).

  **Recommended Agent Profile**:
  - Category: `quick` - Reason: Pure JS data module, no UI, straightforward logic
  - Skills: []

  **Parallelization**: Can Parallel: YES | Wave 1 | Blocks: [4,5,6] | Blocked By: []

  **References**:
  - Placement: `app.js:1-8` - data arrays; insert after line 8, before line 182
  - Pattern: `app.js:293-295` - `escapeHtml` utility function style

  **Acceptance Criteria**:
  - [ ] `menuState` object exists with `minimized`, `buttonEdge`, `buttonTop` properties
  - [ ] `safeStorage.get(key, fallback)` returns parsed JSON or fallback on error
  - [ ] `safeStorage.set(key, value)` writes JSON to localStorage, does not throw
  - [ ] `loadState()` reads `menu-minimized` and `menu-button-pos`, applies defaults if absent
  - [ ] `saveState()` writes current menuState to localStorage

  **QA Scenarios**:
  ```
  Scenario: State persists across reload
    Tool: Playwright
    Steps: (1) Open app, (2) set menuState.minimized=true and saveState(), (3) Reload page, (4) Check menuState.minimized
    Expected: menuState.minimized is true after reload
    Evidence: .sisyphus/evidence/task-2-state-persist.png

  Scenario: localStorage failure handled gracefully
    Tool: Playwright
    Steps: (1) Mock localStorage.setItem to throw, (2) Call saveState(), (3) Check no uncaught error
    Expected: No exception thrown, app continues
    Evidence: .sisyphus/evidence/task-2-error-handling.png
  ```

  **Commit**: YES | Message: `feat(menu): add state management and localStorage helpers` | Files: [app.js]

- [x] 3. Add minimize pill toggle to menu footer

  **What to do**: Add the minimize pill toggle button to the menu footer in `index.html` and render it in `app.js`. In `index.html`, add a `<button class="minimize-pill">` inside the footer-actions div with a chevron-down SVG icon and "Minimize" label text. In `app.js`, add the SVG markup for the chevron-down icon to `iconMarkup` (key: `chevronDown`). Ensure the button is keyboard-accessible with aria-label="Minimize menu". The pill toggle should sit inline with the existing footer actions (Navigate, up/down buttons, divider, Open button).

  **Must NOT do**: Do not add click handler yet (that's Task 4). Do not remove any existing footer elements. Do not change the footer layout structure.

  **Recommended Agent Profile**:
  - Category: `visual-engineering` - Reason: HTML structure + SVG icon, must match existing footer visual style
  - Skills: [`frontend-ui-ux`] - needed for SVG icon and pill styling

  **Parallelization**: Can Parallel: YES | Wave 1 | Blocks: [4,5] | Blocked By: []

  **References**:
  - HTML: `index.html:49-76` - footer structure to insert into
  - HTML: `index.html:66-75` - footer-button-primary pattern for button structure
  - JS: `app.js:49-179` - iconMarkup pattern for SVG definitions
  - CSS: styles.css new `.minimize-pill` rule from Task 1

  **Acceptance Criteria**:
  - [ ] `<button class="minimize-pill">` exists in `index.html` inside the footer-actions div
  - [ ] Button has `aria-label="Minimize menu"` and `type="button"`
  - [ ] Button contains a chevron-down SVG and "Minimize" text label
  - [ ] `iconMarkup.chevronDown` contains the SVG markup string
  - [ ] Pill toggle appears visually in footer between navigate buttons and Open button

  **QA Scenarios**:
  ```
  Scenario: Pill toggle renders in DOM
    Tool: Playwright
    Steps: (1) Open app, (2) Check `.minimize-pill` element exists, (3) Verify it has aria-label "Minimize menu"
    Expected: Button exists in DOM with correct aria-label
    Evidence: .sisyphus/evidence/task-3-pill-renders.png

  Scenario: Pill is keyboard focusable
    Tool: Playwright
    Steps: (1) Open app, (2) Press Tab until pill toggle receives focus, (3) Verify :focus-visible outline visible
    Expected: Pill toggle receives keyboard focus with visible outline
    Evidence: .sisyphus/evidence/task-3-keyboard-focus.png
  ```

  **Commit**: YES | Message: `feat(menu): add minimize pill toggle to footer` | Files: [index.html, app.js]

- [x] 4. Implement minimize/restore toggle behavior

  **What to do**: Wire up the click handler for the minimize pill toggle and the floating button to toggle between minimized and expanded states. On pill click: hide `.command-menu`, create/show `.floating-menu-btn` (if not exists), set `menuState.minimized = true`, call `saveState()`. On floating button click (when not dragged — see Task 5 threshold): hide floating button, show `.command-menu`, set `menuState.minimized = false`, call `saveState()`. The floating button should be created dynamically via `document.createElement` and appended to `document.body`. Apply inline `style.top` and `style.left`/`style.right` based on `menuState.buttonEdge` and `menuState.buttonTop`. On page load: call `loadState()` and apply initial state (show menu or floating button accordingly). Preserve any active search query and results — only toggle visibility, never clear state.

  **Must NOT do**: Do not implement drag behavior (Task 5). Do not clear search input or results when minimizing. Do not remove the command-menu from DOM — toggle visibility only.

  **Recommended Agent Profile**:
  - Category: `quick` - Reason: Event wiring with straightforward show/hide logic, dependent on existing state module
  - Skills: []

  **Parallelization**: Can Parallel: YES | Wave 2 | Blocks: [6] | Blocked By: [1,2,3]

  **References**:
  - JS: `app.js:182-198` - existing event listeners pattern (searchInput, clearButton)
  - JS: `app.js:212-214` - classList.toggle for is-hidden pattern
  - CSS: `styles.css:176-178` - `.is-hidden` utility class
  - CSS: styles.css new `.floating-menu-btn` rule from Task 1
  - State: `app.js` new `menuState`, `loadState()`, `saveState()` from Task 2

  **Acceptance Criteria**:
  - [ ] Clicking minimize pill hides `.command-menu` and shows `.floating-menu-btn`
  - [ ] Clicking floating button hides it and shows `.command-menu`
  - [ ] `menuState.minimized` is true after minimize, false after restore
  - [ ] `saveState()` is called after each toggle
  - [ ] On page load with `menu-minimized === 'true'` in localStorage, floating button shows and menu is hidden
  - [ ] Search input value and results are preserved across minimize/restore
  - [ ] Floating button has correct initial position from `menuState.buttonEdge` and `menuState.buttonTop`

  **QA Scenarios**:
  ```
  Scenario: Minimize and restore cycle
    Tool: Playwright
    Steps: (1) Open app, (2) Click .minimize-pill, (3) Verify .command-menu is hidden and .floating-menu-btn visible, (4) Click .floating-menu-btn, (5) Verify menu visible and button hidden
    Expected: Toggle works both directions
    Evidence: .sisyphus/evidence/task-4-toggle-cycle.png

  Scenario: Search state preserved across minimize
    Tool: Playwright
    Steps: (1) Type "farm" in search, (2) Verify results shown, (3) Click minimize pill, (4) Click floating button to restore, (5) Verify search input still has "farm" and results still shown
    Expected: Search state unchanged after minimize/restore
    Evidence: .sisyphus/evidence/task-4-preserve-search.png

  Scenario: Initial state from localStorage
    Tool: Playwright
    Steps: (1) Set localStorage menu-minimized=true, (2) Reload page, (3) Verify floating button visible and menu hidden
    Expected: App respects persisted minimized state on load
    Evidence: .sisyphus/evidence/task-4-persist-load.png
  ```

  **Commit**: YES | Message: `feat(menu): implement minimize/restore toggle behavior` | Files: [app.js]

- [x] 5. Implement drag behavior with snap-to-edges

  **What to do**: Add pointer event handlers to the floating button for drag-to-reposition with snap-to-edges. Use Pointer Events (`pointerdown`, `pointermove`, `pointerup`) for cross-device support. Track drag state: `isDragging`, `dragStartY`, `dragStartTop`. On `pointerdown`: record start position, set `isDragging = false`. On `pointermove`: if total movement > 5px threshold, set `isDragging = true` and move button by delta (clamp between 0 and viewport height - button height). On `pointerup`: if `isDragging` is false, treat as click (restore — handled by Task 4 click handler). If dragged, snap button to nearest edge: if button center X < viewport width / 2, snap left (set `left: 0`); otherwise snap right (set `right: 0`). Update `menuState.buttonEdge` and `menuState.buttonTop`, call `saveState()`. Prevent text selection during drag (`user-select: none` on body while dragging). Add `event.preventDefault()` on pointermove to prevent scroll interference on touch.

  **Must NOT do**: Do not use mouse events — use Pointer Events only. Do not modify the click handler from Task 4 (drag-vs-click logic must be here: only trigger click when `isDragging === false`). Do not allow dragging outside viewport bounds.

  **Recommended Agent Profile**:
  - Category: `unspecified-high` - Reason: Pointer event handling with viewport math and snap logic, moderate complexity
  - Skills: []

  **Parallelization**: Can Parallel: YES | Wave 2 | Blocks: [6] | Blocked By: [1,2,3]

  **References**:
  - JS: `app.js:182-215` - existing event listener attachment pattern
  - CSS: styles.css new `.floating-menu-btn` (has `touch-action: none`, `pointer-events: auto`)
  - Spec: Metis review - drag-vs-click threshold: 5px movement before treating as drag
  - Spec: Interview - snap to left/right edges based on horizontal midpoint

  **Acceptance Criteria**:
  - [ ] Dragging floating button vertically moves it (pointermove updates top style)
  - [ ] Button stays within viewport (top >= 0, top <= viewportHeight - 48)
  - [ ] On pointerup after drag >5px, button snaps to nearest edge (left or right)
  - [ ] On pointerup with movement <5px, does NOT snap — treated as click (restore)
  - [ ] `menuState.buttonEdge` and `menuState.buttonTop` updated after drag
  - [ ] `saveState()` called after snap
  - [ ] Text not selected during drag (user-select: none applied during drag, removed after)
  - [ ] Works on touch devices (Pointer Events cover this)

  **QA Scenarios**:
  ```
  Scenario: Drag floating button and snap to left edge
    Tool: Playwright
    Steps: (1) Minimize menu (click pill), (2) Drag floating button leftward past viewport midpoint, (3) Release, (4) Verify button snapped to left edge, (5) Check menuState.buttonEdge === 'left'
    Expected: Button snaps to left edge, state updated
    Evidence: .sisyphus/evidence/task-5-drag-snap-left.png

  Scenario: Drag below 5px threshold triggers click not drag
    Tool: Playwright
    Steps: (1) Minimize menu, (2) Click floating button without significant movement, (3) Verify menu restores (not dragged)
    Expected: Small movement treated as click, menu restores
    Evidence: .sisyphus/evidence/task-5-click-threshold.png

  Scenario: Button cannot be dragged off-screen
    Tool: Playwright
    Steps: (1) Minimize menu, (2) Attempt to drag button below viewport bottom, (3) Release, (4) Verify button top <= viewportHeight - 48
    Expected: Button clamped within viewport bounds
    Evidence: .sisyphus/evidence/task-5-viewport-clamp.png
  ```

  **Commit**: YES | Message: `feat(menu): add drag-to-reposition with snap-to-edges` | Files: [app.js]

- [x] 6. Integration: viewport resize handling, keyboard a11y, and final wiring

  **What to do**: Handle edge cases that span multiple systems. (a) Add `resize` event listener that clamps floating button position within new viewport bounds and re-snaps to correct edge. (b) Add keyboard support: floating button must be focusable (native `<button>` should handle this; add explicit `tabindex="0"` for safety), Enter/Space triggers restore. (c) Add `loading` state initialization: on page load, if menuState.minimized is true, set up floating button immediately. (d) Clean up: remove event listeners from floating button when it's hidden to prevent memory leaks (or use a single persistent button instance with visibility toggle). (e) Add CSS `transition` for snap animation (smooth left/right snapping using `var(--transition-fast)`).

  **Must NOT do**: Do not introduce new features beyond these edge cases. Do not add debounce/throttle unnecessarily — resize handler only needs basic clamping.

  **Recommended Agent Profile**:
  - Category: `unspecified-high` - Reason: Cross-cutting concerns spanning event cleanup, resize, and a11y
  - Skills: []

  **Parallelization**: Can Parallel: NO | Wave 3 | Blocks: [] | Blocked By: [4,5]

  **References**:
  - JS: `app.js:182-215` - existing event listener pattern
  - CSS: `styles.css:36` - `--transition-fast` for snap animation
  - Spec: Metis review - viewport resize handling, keyboard accessibility

  **Acceptance Criteria**:
  - [ ] On browser resize while minimized, floating button position clamped to new viewport bounds
  - [ ] Floating button snaps to correct edge after resize (stays on same side)
  - [ ] Floating button has `tabindex="0"` and receives keyboard focus
  - [ ] Pressing Enter or Space on focused floating button restores menu
  - [ ] CSS transition applied to left/right/top changes for smooth snap animation
  - [ ] No memory leaks: event listeners cleaned up appropriately
  - [ ] Page load with `menuState.minimized = true` immediately shows floating button (no flash of menu)

  **QA Scenarios**:
  ```
  Scenario: Resize while minimized keeps button in bounds
    Tool: Playwright
    Steps: (1) Minimize menu, (2) Drag button near bottom of viewport, (3) Resize browser to smaller height, (4) Verify button top <= new viewportHeight - 48
    Expected: Button reclamped after resize
    Evidence: .sisyphus/evidence/task-6-resize-clamp.png

  Scenario: Keyboard restores menu from floating button
    Tool: Playwright
    Steps: (1) Minimize menu, (2) Press Tab until floating button focused, (3) Press Enter, (4) Verify menu restored
    Expected: Enter key restores menu
    Evidence: .sisyphus/evidence/task-6-keyboard-restore.png

  Scenario: No menu flash on load when minimized
    Tool: Playwright
    Steps: (1) Set localStorage menu-minimized=true, (2) Reload page, (3) Verify menu panel not briefly visible before floating button appears
    Expected: Only floating button visible from first paint
    Evidence: .sisyphus/evidence/task-6-no-flash.png
  ```

  **Commit**: YES | Message: `fix(menu): handle resize, keyboard a11y, and initialization edge cases` | Files: [app.js, styles.css]

## Final Verification Wave (MANDATORY — after ALL implementation tasks)
> 4 review agents run in PARALLEL. ALL must APPROVE. Present consolidated results to user and get explicit "okay" before completing.
> **Do NOT auto-proceed after verification. Wait for user's explicit approval before marking work complete.**
> **Never mark F1-F4 as checked before getting user's okay.** Rejection or user feedback → fix → re-run → present again → wait for okay.
- [x] F1. Plan Compliance Audit — oracle
- [x] F2. Code Quality Review — unspecified-high
- [x] F3. Real Manual QA — unspecified-high (+ playwright if UI)
- [x] F4. Scope Fidelity Check — deep

## Commit Strategy
- Task 1 → `style(menu): add floating button and minimize pill CSS`
- Task 2 → `feat(menu): add state management and localStorage helpers`
- Task 3 → `feat(menu): add minimize pill toggle to footer`
- Task 4 → `feat(menu): implement minimize/restore toggle behavior`
- Task 5 → `feat(menu): add drag-to-reposition with snap-to-edges`
- Task 6 → `fix(menu): handle resize, keyboard a11y, and initialization edge cases`

## Success Criteria
1. Pill toggle in footer minimizes the menu to a floating round button with grid icon
2. Floating button is draggable vertically and snaps to left/right screen edges
3. Clicking the floating button restores the full menu
4. Menu state and button position persist across page reloads via localStorage
5. Search query and results are preserved across minimize/restore cycles
6. Floating button stays within viewport bounds on resize
7. Floating button is keyboard accessible (Tab focus, Enter/Space to restore)
8. No external dependencies introduced — pure vanilla HTML/CSS/JS
