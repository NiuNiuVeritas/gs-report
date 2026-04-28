---
name: gs-report
description: Convert Guosen Securities Word research reports (.docx) into WeChat public-account Markdown/HTML using the established gzh template. Use when preparing weekly公众号素材 from Word reports, preserving full body text, adding figure/table screenshot markers, footer metadata, and verification.
---

# GS Report

Convert Guosen-style Word research reports into WeChat-ready Markdown/HTML with strict preservation. This skill is for reports whose public-account draft should follow the existing `gzh` template and the team's manual screenshot workflow.

## Workflow

1. Locate the source `.docx`, the base template if provided, and `assets/law.png`.
2. Run the converter:

```bash
python scripts/convert_gs_report.py --docx <report.docx> --output-dir <folder> --template assets/base-template.md
```

Use optional overrides only when automatic extraction fails:

```bash
python scripts/convert_gs_report.py --docx <report.docx> --publication-date 2026年04月28日 --analyst "张欣慰 S0980520060001" --analyst "彭思宇 S0980521060003"
```

3. Run the verifier before reporting completion:

```bash
python scripts/verify_gs_report.py --docx <report.docx> --markdown <generated.md>
```

4. Report the output paths and verification counts. Do not claim the draft is ready if verification reports missing body paragraphs, missing figure/table markers, unresolved placeholders, missing footer metadata, or missing assets.

## Content Rules

- Fully transfer Word body content from the first `国信研报正文-1.正文一级标题` paragraph to the `免责声明` boundary.
- Do not excerpt, rewrite, summarize down, merge, or delete body paragraphs.
- Preserve explicit Word bold runs in正文 paragraphs as `<strong>...</strong>`.
- Omit Word contents/catalog pages and the full disclaimer/office-address appendix.
- Do not output a title-area author/source line; the WeChat account config handles author metadata.
- Extract the top `报告摘要` from the first-page `核心观点` area and format it with the base template style.
- Treat Word paragraph style IDs `20` and `23` as known `核心观点` summary-body starts; keep converter and verifier in sync when adding more IDs.
- Use Chinese numerals `一、二、三` for summary points.
- Use default body line-height `1.6` with paragraph before/after spacing `0`.
- Keep summary blocks WeChat-editor safe: all summary containers use `box-sizing:border-box;width:100%;max-width:100%`, and summary bullets are rendered as `·` marker paragraphs rather than `<ul>/<li>` lists.
- Set the `总结` section's top margin to `0`.

## Figure And Table Rules

- The team screenshots figures/tables directly from Word.
- In generated material, include only the figure/table title positioning line.
- Do not include duplicate source-note positioning lines; source notes come from the Word screenshot.
- Do not add special locator styling such as extra bold, italic, slashes, or brackets.

## Footer Rules

Use this footer order:

1. `注：本文选自国信证券于{发布日期}发布的研究报告《{报告标题}》`
2. One line per analyst: `分析师：{姓名} {证书号}`
3. The standard risk prompt.
4. The `量化藏经阁` WeChat profile card.
5. `law.png` legal declaration image.

The converter should auto-extract publication date from the first-page header and analysts from the first-page analyst block. If either is missing, rerun with `--publication-date` or `--analyst`.

## Resources

- `scripts/convert_gs_report.py`: deterministic converter.
- `scripts/verify_gs_report.py`: completeness verifier.
- `assets/base-template.md`: current base WeChat Markdown style.
- `assets/law.png`: legal declaration image copied into each output asset folder.
- `references/requirements.md`: detailed conversion contract.
