import streamlit as st
from generate_award_slides import build_deck
from io import BytesIO

st.set_page_config(page_title='Award Slide Generator', layout='wide')
st.title('Award Slide Generator')

st.sidebar.title('Instructions')
st.sidebar.markdown("""
Excel columns:
Award Category | Nominee Name | Winner Name | Zone | Placeholder X | Placeholder Y | Placeholder Z

Important:
- Keep nominee and winner in separate rows.
- Winner slides appear after nominee slides for the same category.
- The template background is preserved.
- Slide role is detected by placeholders.
- All placeholders are optional.
""")

excel = st.file_uploader('Upload Excel', type=['xlsx'])
template = st.file_uploader('Upload PowerPoint template', type=['pptx'])

if excel is not None and template is not None and st.button('Generate'):
    try:
        deck = build_deck(excel, template)
        bio = BytesIO()
        deck.save(bio)
        st.success(f'Generated {len(deck.slides)} slide(s).')
        st.download_button('Download PPTX', bio.getvalue(), file_name='Award_Show_Deck.pptx')
    except Exception as e:
        st.error(f'Generation failed: {e}')
