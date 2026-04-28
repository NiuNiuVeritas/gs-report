# GS Report Conversion Contract

## Inputs

- Guosen Securities research report `.docx`.
- Optional base template, normally `assets/base-template.md`.
- Optional publication date override.
- Optional analyst override lines.
- Static footer image `assets/law.png`.

## Outputs

- WeChat Markdown with inline HTML.
- Standalone HTML preview.
- Output asset folder containing `law.png`.
- Verification output.

## Strictness

The default mode is complete transfer, not editorial rewriting. Preserve every Word body paragraph before `免责声明`. Exclude directory/catalog content and the full disclaimer/office-address appendix.

Preserve explicit Word bold runs in body paragraphs as `<strong>...</strong>`.

Preserve Word body bullets. If a Word body paragraph has numbering/bullet properties, render it as a native `<li>` bullet in the generated material and verify expected/actual body bullet counts.

Do not output the author/source metadata row under the title; WeChat account settings handle author metadata.

For appendix section titles, remove leading labels such as `附录一` and `附录二` but keep the remaining title text.

## Summary

Use the first-page `核心观点` area. Preserve its meaning and order while fitting the existing template. Use Chinese numerals. Treat Word paragraph style IDs `20` and `23` as known summary-body starts; keep the converter and verifier style ID lists in sync.

Render Word bullet paragraphs as native `<li>` bullets with controlled inline width styles so copied WeChat drafts retain browser/WeChat `li::marker` behavior.

## Text Spacing

Default line-height is `1.6`. Paragraph before/after spacing is `0`.

## Figures And Tables

Only output title marker lines for figures/tables. These lines are temporary locators for the team to replace with Word screenshots. Do not output source-note lines as separate text.

## Footer

Footer order is source note, analysts, blank final risk prompt, `law.png`. Do not include the WeChat profile card in generated HTML.

Ask the user if publication date, analyst name, analyst certificate number, or `law.png` is missing.
