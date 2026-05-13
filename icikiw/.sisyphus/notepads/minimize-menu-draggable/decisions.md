# Decisions
## minimize-menu-draggable

### Architecture Decisions
- **No external dependencies**: Pure vanilla JS/CSS — consistent with repo AGENTS.md
- **Floating button positioning**: `position: fixed` with JS-managed inline `top`/`left`/`right` styles (not CSS classes)
- **State persistence**: localStorage keys `menu-minimized` (boolean string) and `menu-button-pos` (JSON `{edge, top}`)
- **Drag events**: Pointer Events API instead of mouse/touch events — covers all input types
- **Visibility toggle**: Use `.is-hidden` class (existing pattern) rather than removing DOM elements
- **Search preservation**: Only toggle panel visibility — never clear `searchInput.value` or re-render

### Constants
- Floating button size: 48px (width and height)
- Z-index: 9999
- Drag threshold: 5px (movement below this = click, above = drag)
- Default position: right edge, top 200px
- localStorage keys: `menu-minimized`, `menu-button-pos`

- 2026-05-12 code-quality re-review: APPROVE. Pointercancel cleanup is present at lines 373-376; focus-visible rgba values match the existing stylesheet pattern; safeStorage wraps localStorage access in try/catch; no obvious dead code, unused vars, or global leaks found; `node --check app.js` passed.
