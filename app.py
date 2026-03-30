import streamlit as st
import tempfile
import os
from pathlib import Path
from excel_to_ubl_xml import build_xml

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Excel → UBL XML",
    page_icon="📄",
    layout="centered",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}

/* Background */
.stApp {
    background-color: #0f0f0f;
    color: #e8e8e8;
}

/* Hide default Streamlit header/footer chrome */
#MainMenu, footer, header {visibility: hidden;}

/* Main container */
.block-container {
    max-width: 640px;
    padding-top: 4rem;
    padding-bottom: 4rem;
}

/* Title area */
.app-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.1rem;
    font-weight: 600;
    letter-spacing: 0.15em;
    text-transform: uppercase;
    color: #c8f55a;
    margin-bottom: 0.25rem;
}
.app-subtitle {
    font-size: 0.85rem;
    color: #555;
    font-family: 'IBM Plex Mono', monospace;
    letter-spacing: 0.05em;
    margin-bottom: 3rem;
}

/* File uploader */
[data-testid="stFileUploader"] {
    background: #1a1a1a;
    border: 1.5px dashed #2e2e2e;
    border-radius: 4px;
    padding: 1rem;
    transition: border-color 0.2s;
}
[data-testid="stFileUploader"]:hover {
    border-color: #c8f55a;
}
[data-testid="stFileUploader"] label {
    color: #888 !important;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.82rem;
}

/* Button */
.stButton > button {
    background: #c8f55a;
    color: #0f0f0f;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 600;
    font-size: 0.85rem;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    border: none;
    border-radius: 3px;
    padding: 0.65rem 2rem;
    width: 100%;
    margin-top: 1rem;
    cursor: pointer;
    transition: background 0.15s, transform 0.1s;
}
.stButton > button:hover {
    background: #d4ff66;
    transform: translateY(-1px);
}
.stButton > button:active {
    transform: translateY(0);
}

/* Download button */
.stDownloadButton > button {
    background: transparent;
    color: #c8f55a;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 600;
    font-size: 0.85rem;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    border: 1.5px solid #c8f55a;
    border-radius: 3px;
    padding: 0.65rem 2rem;
    width: 100%;
    margin-top: 0.5rem;
    cursor: pointer;
    transition: all 0.15s;
}
.stDownloadButton > button:hover {
    background: #c8f55a;
    color: #0f0f0f;
}

/* Success / error boxes */
.stSuccess {
    background: #1a2a0a !important;
    border: 1px solid #4a7a10 !important;
    border-radius: 3px !important;
    color: #c8f55a !important;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.82rem;
}
.stError {
    background: #2a0a0a !important;
    border: 1px solid #7a1010 !important;
    border-radius: 3px !important;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.82rem;
}

/* Divider */
hr {
    border-color: #1e1e1e;
    margin: 2rem 0;
}

/* Info text */
.info-row {
    display: flex;
    justify-content: space-between;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.75rem;
    color: #444;
    margin-top: 3rem;
    padding-top: 1rem;
    border-top: 1px solid #1e1e1e;
}
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown('<div class="app-title">Invoice Converter</div>', unsafe_allow_html=True)
st.markdown('<div class="app-subtitle">xlsx → ubl xml · serbian e-faktura</div>', unsafe_allow_html=True)

# ── Upload ────────────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Drop your Excel file here or click to browse",
    type=["xlsx"],
    label_visibility="visible",
)

# ── Convert ───────────────────────────────────────────────────────────────────
if uploaded:
    st.markdown(f"**`{uploaded.name}`** — ready to convert")

    if st.button("Convert to XML"):
        with st.spinner("Processing..."):
            try:
                with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                    tmp.write(uploaded.read())
                    tmp_path = tmp.name

                xml_bytes = build_xml(tmp_path)
                os.unlink(tmp_path)

                output_name = Path(uploaded.name).stem + ".xml"

                st.success(f"✓ Converted successfully — {len(xml_bytes):,} bytes")

                st.download_button(
                    label="⬇  Download XML",
                    data=xml_bytes,
                    file_name=output_name,
                    mime="application/xml",
                )

            except Exception as e:
                try:
                    os.unlink(tmp_path)
                except Exception:
                    pass
                st.error(f"Conversion failed:\n\n{e}")

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="info-row">
    <span>Reads: General · Invoice Lines · Totals · Invoicing</span>
    <span>Schema: EN 16931 / mfin.gov.rs 2022</span>
</div>
""", unsafe_allow_html=True)
