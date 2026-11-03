# H&H Slide Studio

A static React/Vite web app for turning faculty source material into editable H&H-branded teaching PowerPoint decks. Designed for deployment to Cloudflare Pages.

## Local development

```bash
npm install
npm run dev
```

## Production build

```bash
npm run build
```

Deploy the generated `dist/` directory to Cloudflare Pages. Build command: `npm run build`; output directory: `dist`.

## Staff quick guide

1. Open **H&H Slide Studio** and choose Lecture, Quiz or Workshop.
2. Drop a PDF, DOCX, TXT or Markdown file into **Start with your content**. Alternatively, add slides manually from the deck editor.
3. Enter the course or unit name, then choose **Build slide outline**.
4. Select each slide thumbnail to edit its heading, teaching text and layout. Put each bullet or quiz option on its own line.
5. Add or remove slides as needed. Review the complete outline before export.
6. Choose **Export PowerPoint**. The output is a widescreen, editable `.pptx` with a branded Slide Master and H&H theme settings.

## Brand system

The supplied brand sheet specifies H&H teal `#075766`, magenta `#A1164F`, navy `#24205F`, cyan `#2CB8CE`, gold `#F5C84C`, paper `#F7F5F0`, pale `#EDF5F6`, and charcoal `#263238`. Headings use Manrope; body text uses Source Sans 3. The app uses these values in the exported theme and master slide.

## Current scope and limitations

- Static client-side app; no server, API key or database required.
- Source import currently supports PDF, DOCX, TXT and MD. PPTX import is intentionally not enabled; convert source slides to PDF or DOCX first.
- Text extraction and slide grouping are heuristic, not generative AI. Staff should review and edit the generated outline before presenting.
- The generated deck uses editable PowerPoint text and shapes, a reusable master, and a theme. The brand sheet provided was a single composite image rather than individual vector logo files; the exported cover therefore uses a clean typographic H&H lock-up. Replace it with the official vector logo once supplied.
- Accessibility: use high-contrast text, meaningful slide titles, readable font sizes and descriptive alt text for any future image assets. Automated contrast checks and full accessibility auditing are not yet implemented.

## Brand asset handoff needed for final brand sign-off

The supplied reference is a raster brand-board composite, not a separate production logo file. The prototype extracts a small transparent logo for the web UI and uses a typographic lock-up in exported slides. For final institutional rollout, replace this with the official SVG/transparent PNG primary logo and confirm the logo clear-space measurement, minimum size, approved photo library, and exact accessibility/contrast rules with the brand owner. Fonts are referenced by family name in the PowerPoint; they are not embedded, so install Manrope and Source Sans 3 on staff computers to avoid substitution.
