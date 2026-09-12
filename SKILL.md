---
name: wechat-article-word-archiver
description: Archive one or more WeChat Official Account articles into separate illustrated Word documents, each with its own folder and image subfolder, then validate and package the result. Use when a user supplies mp.weixin.qq.com article URLs and asks for 带图 Word、每篇单独文件夹、公众号推文归档、批量打包 or a ZIP deliverable.
---

# WeChat Article Word Archiver

Create faithful, reviewable archives of public WeChat articles. Treat the source pages as authoritative; never invent missing text, images, titles, or item numbers.

## Inputs that change on every run

Ask only for missing required values. Accept either a pasted list or the JSON schema in [references/input-format.md](references/input-format.md).

- `archive_name`: output package name.
- `articles`: ordered items, each with `id`, `label`, and `url`.
- Optional `html_path`: a previously saved HTML file when direct retrieval is unavailable.
- Optional output parent. Default to the current workspace's `outputs` directory.

The article count, identifiers, labels, URLs, output path, and source content are variables. Do not copy examples into a live run.

## Fixed output contract

For every article create exactly:

```text
<output-parent>/<archive-name>/
  <id>-<safe-label>/
    <id>-<safe-label>.docx
    图片/
      图片01.jpg
      ...
  归档清单.json
<output-parent>/<archive-name>.zip
```

The Word file must contain the article text and every successfully archived static image. The `图片` folder must contain the same image set as separate files. Keep the input order and preserve non-consecutive identifiers; never silently renumber a missing item.

## Required workflow

1. Parse the supplied list into the JSON input schema. Preserve labels and identifiers exactly, except replace Windows-invalid filename characters.
2. Load bundled document dependencies with `load_workspace_dependencies`. Use the bundled Python when available.
3. Retrieve each source read-only. For WeChat query URLs, normalize a lone `biz` key to `__biz`; keep short `/s/<token>` URLs intact. Use a browser-like user agent and preserve the original user URL in the document.
4. Validate the fetched HTML before building anything: require a non-empty `#activity-name` title and `#js_content` body. CAPTCHA, safety-warning, parameter-error, login, or empty pages are not article content.
5. If direct retrieval fails, use an already-authorized browser/Computer Use session to open the exact user URL. Do not bypass platform checks. If the article still cannot be read, request a saved HTML/PDF, copied full article, or continuous screenshots and mark that item blocked rather than fabricating a partial archive.
6. Run `scripts/archive_wechat_articles.py` with the JSON input and output parent. This downloads and normalizes images, writes each DOCX, performs structural checks, writes `归档清单.json`, and creates the ZIP.
7. Render every generated DOCX. First use the documents skill renderer if its bundled LibreOffice runtime works. On Windows, if `soffice.exe` is unavailable, use `scripts/render_docx_windows.ps1` with Microsoft Word COM and bundled Poppler `pdftoppm`.
8. Run `scripts/make_contact_sheets.py` over the rendered page images. Inspect every contact sheet for blank pages, clipped text, misplaced images, broken glyphs, error pages, or unexpected large whitespace.
9. Re-run structural validation after any repair. Deliver only when every requested item passes.

## Build command

```powershell
<bundled-python> scripts/archive_wechat_articles.py `
  --input <input-json> `
  --output-parent <output-parent> `
  --cache-dir <temporary-cache> `
  --zip
```

Use a new cache directory per run. Add `--overwrite` only when replacing the exact archive directory named by `archive_name` is clearly intended.

## Document rules

- Letter portrait; readable margins; Microsoft YaHei for Chinese and Latin text.
- Centered black title, source metadata when present, and a clickable `原始网页` link.
- Readable body spacing; centered images scaled within the printable page.
- Extract text and images in document order from `#js_content`.
- Prefer image URLs in `data-src`, `data-original`, `data-echo`, then `src`.
- Skip GIFs, data/blob URLs, failed downloads, files smaller than 700 bytes, and images below 80×60 pixels. Normalize accepted images to JPEG or PNG through Pillow so Word renders them reliably.
- Omit WeChat chrome such as like/share/comment controls. Do not add summaries, captions, or claims absent from the source.
- Dynamic video/audio/mini-program content is not equivalent to a static archive. State any such limitation in the manifest and handoff.

## Validation gate

Do not report completion until all checks pass:

- requested item count equals folder count;
- every article folder contains exactly one DOCX and one `图片` directory;
- each DOCX has non-empty paragraphs;
- `python-docx` inline-shape count equals the number of archived image files;
- every DOCX renders successfully and rendered page count is greater than zero;
- contact sheets cover every rendered page and visual inspection finds no broken or blank output;
- `归档清单.json` lists every input identifier, source URL, document name, text-block count, archived image count, embedded image count, and validation state;
- ZIP exists and contains every DOCX, image, and manifest entry.

If one item fails, identify that item and the failed check. Do not present the package as fully complete.

## Handoff format

Report the archive folder, ZIP, item count, image count, and rendered-page count. Link the ZIP and each final DOCX exactly once. Mention blocked or omitted dynamic elements explicitly. Temporary HTML, PDFs, rendered PNGs, and contact sheets are QA artifacts and are not part of the deliverable unless requested.

