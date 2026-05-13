# AGENTS.md

## Repo shape
- Static frontend only: `index.html` loads `app.js` and `styles.css`.
- There is no package manifest, build tool, test runner, or CI config in this repo.
- The app is a single-page quick-action menu; keep changes consistent with that small surface area.

## What to edit
- `app.js`: menu data, filtering, state, event handling.
- `styles.css`: all visual changes.
- `index.html`: structure and accessibility hooks only; keep IDs/classes stable unless JS is updated with them.

## Working rules
- Open the app directly in a browser to verify changes; there is no repo command to run.
- Prefer minimal, local edits. Avoid adding tooling or architecture that the repo does not already use.
- Preserve the existing plain-vanilla HTML/CSS/JS style unless the user explicitly asks for a rewrite.

## Existing instructions
- No repo-local `CLAUDE.md`, `.cursorrules`, `.cursor/`, `copilot-instructions`, or `opencode.json` were present when this file was written.
