H&H Slide Studio — brand-board sample workflow update

Apply these files at the root of the existing H&H Slide Studio repository, preserving folder paths:
- src/main.jsx
- src/hh-brand-refresh.css
- public/assets/hh-brand-guidelines.jpg
- package.json

The existing src/style.css and public/assets/hh-logo.png are expected to remain in the repository.

After uploading/committing the files:
1. Allow Cloudflare Pages to rebuild the site.
2. Open Presentation builder.
3. Click “Load H&H brand-board sample” (or upload the supplied H&H brand guidelines image).
4. Review the editable brand summary slides and the final original-image reference slide.
5. Click “Export PowerPoint”.

The sample deck contains editable text and shapes for the brand promise, logo use, colour palette, typography, iconography, learning pathway and brand applications. The final slide preserves the supplied brand board as an image reference. Image text is not OCR-extracted automatically; the editable summaries are curated from the supplied board.

Note: JSX and CSS syntax checks passed. A full Vite production build and browser export test were not run in this environment because the complete project dependencies/runtime were not installed here. Verify the Cloudflare deployment and open the exported PPTX in PowerPoint before treating the update as production-ready.
