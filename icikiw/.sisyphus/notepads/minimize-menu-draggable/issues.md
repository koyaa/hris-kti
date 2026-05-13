# Issues & Gotchas
## minimize-menu-draggable

### Known Risks
- **localStorage in private browsing**: Can throw on setItem — handled via safeStorage try/catch
- **Viewport resize while minimized**: Button may go off-screen — handled via resize listener with clamping
- **Drag vs click confusion**: Small movements during click could trigger drag — handled via 5px threshold
- **Touch scrolling on mobile**: pointermove may conflict with page scroll — handled via touch-action: none and preventDefault
- **Flash of menu on load**: If minimized state, menu should not briefly appear — handled via early state check before initial render
### Task 4 Verification Blocker
- Playwright browser verification could not run because the environment lacks a Chrome distribution (`/opt/google/chrome/chrome` missing). An install attempt hit a sudo/password requirement in this container.
