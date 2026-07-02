# Award Slide Generator

## Excel format
Use this header row:
Award Category | Nominee Name | Winner Name | Zone | Placeholder X | Placeholder Y | Placeholder Z

Rules:
- `Nominee Name` is for nominee rows.
- `Winner Name` is for winner rows.
- If a category has both nominees and a winner, keep them as separate rows.
- `Zone` is optional but recommended.
- `Placeholder X/Y/Z` are reserved for future use.

## PowerPoint template
Upload a PPTX with two slides if you want both nominee and winner outputs:
- Nominee slide: contains `<<NOMINEES>>`
- Winner slide: contains `<<WINNER>>`

All placeholders are optional.

Supported tokens:
- `<<ZONE>>`
- `<<AWARD CATEGORY>>`
- `<<NOMINEES>>`
- `<<WINNER>>`
- `<<nominees-word>>` -> literal `NOMINEES`
- `<<winner-word>>` -> literal `WINNER`
- `<<PLACEHOLDER X>>`
- `<<PLACEHOLDER Y>>`
- `<<PLACEHOLDER Z>>`

## Slide behavior
- Nominee slides are grouped by `Zone + Award Category`.
- Winner slides are grouped by `Zone + Award Category`.
- Winner slides appear after nominee slides for the same category.
- The template background/design is preserved by cloning the matching template slide.