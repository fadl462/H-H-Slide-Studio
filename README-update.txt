H&H Slide Studio — sample loading and PowerPoint export reliability fix

Replace these files in the root of the existing GitHub repository, preserving paths:
- src/main.jsx
- src/hh-brand-refresh.css
- package.json

IMPORTANT: The previous update already added the following assets to the repository. Keep them in place; do not delete or rename them:
- public/assets/hh-brand-guidelines.jpg
- public/assets/hh-logo.png
- public/assets/hh-sample-photos/ (13 JPEG files)

What this update changes:
1. The H&H sample deck loads automatically when the builder opens. This prevents accidentally exporting the six-slide placeholder deck while intending to test the brand-board sample.
2. The sample deck contains 11 slides, including two photo galleries and the original brand board. The 12 gallery photos are embedded in the presentation data.
3. The editor now displays real photo thumbnails on gallery slides, instead of showing a generic text-only preview.
4. The export no longer rewrites the generated PowerPoint ZIP package with JSZip. It downloads the PPTX blob produced directly by PptxGenJS after checking the ZIP signature and the presence of all gallery photos. This removes the extra post-processing step that could create Office compatibility problems.
5. H&H colours are applied directly to editable slide elements; Manrope and Source Sans 3 remain the presentation theme fonts. The exported slide content remains editable. The custom colour picker palette is not injected into the Office theme in this safer export path.
6. If bundled sample assets fail to load, the editor clears the placeholder deck and displays the actual asset error rather than allowing a misleading export.

Install:
1. Replace the three files above in GitHub.
2. Keep the existing public/assets photo files and logo.
3. Commit and push the changes, then wait for Cloudflare Pages to finish deploying.
4. Hard-refresh the deployed page (Ctrl+F5). The 11-slide H&H sample should load automatically.
5. Select the two gallery slides and confirm that the photos appear in the preview.
6. Export the PowerPoint and open the downloaded .pptx in Microsoft PowerPoint.

Validation performed on the updated source:
- JSX transpilation: passed.
- CSS parsing: passed.
- Export code now validates the PPTX ZIP signature and checks that gallery photos are present before triggering download.

Limitations of this validation:
- A full Vite production build and a browser-driven export from the live Cloudflare deployment were not run in this environment. Confirm the live deployment after pushing the update.
