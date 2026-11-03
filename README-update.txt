H&H Slide Studio — embedded photography and PowerPoint export repair

Apply these files at the root of the existing H&H Slide Studio repository, preserving folder paths:
- src/main.jsx
- src/hh-brand-refresh.css
- public/assets/hh-brand-guidelines.jpg
- public/assets/hh-sample-photos/ (all 13 JPEG files)
- package.json

Keep the existing src/style.css and public/assets/hh-logo.png in place. Do not delete or replace them.

What this update fixes:
1. The H&H sample deck now embeds the source photography in the presentation itself: the cover, course photography gallery, brand-application gallery and original-board reference slide.
2. The gallery preserves each photo's aspect ratio instead of stretching it into a fixed box.
3. The PowerPoint theme patch now converts the generated Blob to an ArrayBuffer before JSZip reads it. This avoids the Blob-read incompatibility that could produce a downloaded file PowerPoint could not open.
4. The exporter checks that essential PowerPoint package parts exist before downloading.

Deployment and test:
1. Upload/replace the listed files in the repository, preserving the folder paths.
2. Commit and push; allow Cloudflare Pages to complete its build.
3. Open Presentation builder and click “Load H&H brand-board sample”.
4. Confirm the course and brand-application gallery slides show the people photographs.
5. Export the PowerPoint, then open it in Microsoft PowerPoint or LibreOffice Impress.

Validation performed on this update: JSX transpilation passed; the PPTX generation and theme-patching sequence was tested with embedded images; the generated PPTX passed ZIP integrity checks and opened in LibreOffice Impress. A full Cloudflare production deployment was not performed from this environment.
