# Award Slide Generator

## Files
- generate_award_slides.py — core logic (Excel parsing, slide duplication, template preservation)
- streamlit_app.py — web UI
- requirements.txt — dependencies

## Excel format
Award Category | Nominee Name | Winner Name | Zone | Placeholder X | Placeholder Y | Placeholder Z

- Nominee rows use Nominee Name.
- Winner rows use Winner Name.
- Keep nominee and winner rows separate.
- Zone is optional but recommended.

## PowerPoint template
Upload a PPTX with:
- a nominee slide containing <<Nominee Name>> (or <<NOMINEES>>)
- a winner slide containing <<Winner Name>> (or <<WINNER>>)

Supported tokens:
<<ZONE>>, <<AWARD CATEGORY>>, <<NOMINEES>>, <<WINNER>>, <<nominees-word>>, <<winner-word>>,
<<PLACEHOLDER X>>, <<PLACEHOLDER Y>>, <<PLACEHOLDER Z>>

## Behavior
- Winner slides appear after nominee slides for the same Zone + Category.
- The actual template slide (background, images, design) is duplicated at the XML level,
  so the original design is preserved exactly — not reconstructed from scratch.

## Run locally
pip install -r requirements.txt
streamlit run streamlit_app.py
