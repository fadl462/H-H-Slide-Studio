H&H Slide Studio — Restore Sample Deck button update

Replace this file in the existing GitHub repository, preserving its path:
- src/main.jsx

What changed:
1. Renamed the sample control to “Restore H&H Sample Deck”.
2. Tracks whether the current deck has been edited, imported, or rebuilt.
3. If the current deck has changes, clicking Restore asks for confirmation before replacing the current slides and discarding those edits.
4. If the deck has not been changed, the sample reloads immediately without an unnecessary confirmation.
5. Successfully restoring the sample clears the modified state so the next restore does not ask again unless the deck is edited.

Install:
1. Replace src/main.jsx in GitHub with the file in this ZIP.
2. Do not remove or rename the existing H&H logo, brand-board image, or sample photo assets.
3. Commit and push the change, then wait for Cloudflare Pages to deploy.
4. Hard-refresh the site (Ctrl+F5).
5. Confirm the button reads “Restore H&H Sample Deck”. Edit a slide and click it to verify the confirmation appears. Cancel should preserve the edits; confirm should reload the original 11-slide sample.

Validation:
- JSX transpilation passed.
- ZIP archive integrity passed.
- Live Cloudflare deployment has not been tested from this update package.
