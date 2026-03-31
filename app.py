import streamlit as st
import tempfile
import os
import zipfile
import io
import re
from docx2pdf import convert


def convert_word_to_pdf(uploaded_files, progress_bar, status_text):
    """Convert uploaded Word files to PDFs and return them as a zip."""
    pdf_files = []
    total = len(uploaded_files)

    with tempfile.TemporaryDirectory() as tmp_dir:
        input_dir = os.path.join(tmp_dir, "input")
        output_dir = os.path.join(tmp_dir, "output")
        os.makedirs(input_dir)
        os.makedirs(output_dir)

        # Save uploaded files to temp input directory
        for uploaded_file in uploaded_files:
            file_path = os.path.join(input_dir, uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

        # Convert each file individually for progress tracking
        for i, uploaded_file in enumerate(uploaded_files):
            file_name = uploaded_file.name
            base_name = os.path.splitext(file_name)[0]
            input_path = os.path.join(input_dir, file_name)
            output_path = os.path.join(output_dir, f"{base_name}.pdf")

            status_text.markdown(f"<span style='color:#cbd5e1;'>Converting: {file_name}</span>", unsafe_allow_html=True)
            try:
                convert(input_path, output_path)
                with open(output_path, "rb") as f:
                    pdf_files.append((f"{base_name}.pdf", f.read()))
            except Exception as e:
                st.error(f"Failed to convert **{file_name}**: {e}")

            progress_bar.progress((i + 1) / total)

        status_text.markdown("<span style='color:#38ef7d; font-weight:600;'>All conversions complete!</span>", unsafe_allow_html=True)

    return pdf_files


def create_zip(pdf_files):
    """Bundle all PDF files into a single ZIP archive."""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in pdf_files:
            zf.writestr(name, data)
    zip_buffer.seek(0)
    return zip_buffer


def format_size(size_bytes):
    """Format bytes to human-readable size."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.2f} MB"


# ─── Page Configuration ──────────────────────────────────────────────
st.set_page_config(
    page_title="W2PDF",
    page_icon="",
    layout="wide",  # Changed to wide for better side-by-side view
)

# ─── Custom CSS ───────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

/* Custom premium app background */
.stApp {
    background-color: #0f172a !important;
    background-image: 
        radial-gradient(at 0% 0%, rgba(102, 126, 234, 0.15) 0px, transparent 50%),
        radial-gradient(at 100% 100%, rgba(118, 75, 162, 0.15) 0px, transparent 50%) !important;
}

/* Hide Streamlit default UI elements */
#MainMenu, footer, header { visibility: hidden; }

/* Reduce overall top gap */
.block-container {
    padding-top: 1.5rem !important;
    padding-bottom: 2rem !important;
}

/* Header gradient */
.main-header {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    padding: 2rem 2rem;
    border-radius: 16px;
    text-align: center;
    margin-bottom: 2rem;
    box-shadow: 0 8px 32px rgba(102, 126, 234, 0.25);
}
.main-header h1 {
    color: white;
    font-size: 2.2rem;
    font-weight: 700;
    margin: 0;
}
.main-header p {
    color: rgba(255,255,255,0.85);
    font-size: 1rem;
    margin: 0.5rem 0 0 0;
}

/* Upload wrapper */
[data-testid="stFileUploader"] {
    border: 1px solid rgba(102, 126, 234, 0.4);
    border-radius: 12px;
    padding: 0.8rem;
    background: rgba(30, 41, 59, 0.6);
    transition: border-color 0.3s ease;
}
[data-testid="stFileUploader"]:hover {
    border-color: #764ba2;
}

/* Internal dropzone dark mode override */
[data-testid="stFileUploaderDropzone"] {
    background-color: transparent !important;
}
[data-testid="stFileUploaderDropzone"] * {
    color: #cbd5e1 !important;
}
[data-testid="stFileUploaderDropzone"] button {
    background: rgba(102, 126, 234, 0.15) !important;
    color: #f8fafc !important;
    border: 1px solid rgba(102, 126, 234, 0.4) !important;
    border-radius: 8px !important;
}

/* Prevent text truncation inside uploader limits */
/* Compact uploader text to fit cleanly without stretching height */
[data-testid="stFileUploader"] small {
    white-space: nowrap !important;
    font-size: 0.75rem !important;
    overflow: visible !important;
    text-overflow: clip !important;
}

/* Force dark mode typography globally (in case config.toml requires server restart) */
h1, h2, h3, h4, h5, h6 {
    color: #f8fafc !important;
}
.stMarkdown p {
    color: #cbd5e1 !important;
}
label, [data-testid="stWidgetLabel"] p, .stFileUploader label p {
    color: #f8fafc !important;
}

/* Explicitly fix Expander Header and File Card typography */
[data-testid="stExpander"] summary p, [data-testid="stExpander"] p, .stAlert p {
    color: #f8fafc !important;
}
.file-card strong {
    color: #f8fafc !important;
}

/* Force dark expander background */
[data-testid="stExpander"] {
    background-color: rgba(30, 41, 59, 0.6) !important;
    border: 1px solid rgba(102, 126, 234, 0.3) !important;
    border-radius: 12px !important;
}
[data-testid="stExpander"] details {
    background-color: transparent !important;
    border: none !important;
}
[data-testid="stExpander"] summary {
    background-color: rgba(30, 41, 59, 0.8) !important;
    border-radius: 12px !important;
}

/* Fix uploaded file items in the left column */
[data-testid="stFileUploader"] [data-testid="stMarkdownContainer"] p {
    color: #e2e8f0 !important;
}
[data-testid="stFileUploader"] small {
    color: #94a3b8 !important;
}

/* Buttons */
.stButton > button {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white;
    border: none;
    border-radius: 10px;
    padding: 0.7rem 2rem;
    font-weight: 600;
    font-size: 1rem;
    width: 100%;
    transition: all 0.3s ease;
    box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
}
.stButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(102, 126, 234, 0.45);
}

/* Download button */
.stDownloadButton > button {
    background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
    color: white;
    border: none;
    border-radius: 10px;
    padding: 0.7rem 2rem;
    font-weight: 600;
    font-size: 1rem;
    width: 100%;
    transition: all 0.3s ease;
    box-shadow: 0 4px 15px rgba(17, 153, 142, 0.3);
}
.stDownloadButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(17, 153, 142, 0.45);
}

/* File list cards */
.file-card {
    background: rgba(30, 41, 59, 0.7);
    border-left: 4px solid #667eea;
    border-radius: 8px;
    padding: 0.6rem 1rem;
    margin-bottom: 0.4rem;
    font-size: 0.9rem;
    transition: background 0.2s ease;
}
.file-card:hover {
    background: rgba(51, 65, 85, 0.9);
}

/* Stats badges */
.stat-badge {
    display: inline-block;
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white;
    padding: 0.35rem 1rem;
    border-radius: 20px;
    font-size: 0.85rem;
    font-weight: 600;
    margin-right: 0.5rem;
}

/* Success container */
.success-box {
    background: linear-gradient(135deg, rgba(17, 153, 142, 0.15), rgba(56, 239, 125, 0.10));
    border: 1px solid rgba(17, 153, 142, 0.4);
    border-radius: 12px;
    padding: 1.5rem;
    margin: 1rem 0;
    text-align: center;
}
.success-box h3 {
    color: #38ef7d !important;
}
.success-box p {
    color: #cbd5e1 !important;
}

/* Force ALL remaining text elements to be light */
span, div, p, li {
    color: #e2e8f0;
}

/* Uploaded file items in the native Streamlit uploader */
[data-testid="stFileUploaderFile"] {
    color: #e2e8f0 !important;
}
[data-testid="stFileUploaderFile"] * {
    color: #e2e8f0 !important;
}
[data-testid="stFileUploaderFile"] small {
    color: #94a3b8 !important;
}

/* Pagination arrows in file uploader */
[data-testid="stFileUploader"] button[kind="icon"] svg,
[data-testid="stFileUploader"] button svg,
[data-testid="stFileUploader"] [data-testid="baseButton-headerNoPadding"] svg {
    fill: #f8fafc !important;
    stroke: #f8fafc !important;
    color: #f8fafc !important;
}
[data-testid="stFileUploader"] button {
    color: #f8fafc !important;
}

/* Status text ("All conversions complete!") */
[data-testid="stText"] {
    color: #38ef7d !important;
}

/* Hide fullscreen button on images */
button[title="View fullscreen"],
[data-testid="StyledFullScreenButton"] {
    display: none !important;
}

</style>
""", unsafe_allow_html=True)

# ─── Header ──────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>W2PDF - Word files to PDF Bulk Converter</h1>
    <p>Upload your Word documents and download them all as PDFs instantly</p>
</div>
""", unsafe_allow_html=True)

# ─── Layout Split (Side-by-Side) ──────────────────────────────────────
left_col, right_col = st.columns([1.1, 1], gap="large")

with left_col:
    st.subheader("Upload Documents")
    uploaded_files = st.file_uploader(
        "Drop your Word files here",
        type=["docx", "doc"],
        accept_multiple_files=True,
        help="Select one or more .docx / .doc files to convert to PDF",
    )

    if uploaded_files:
        total_size = sum(f.size for f in uploaded_files)
        formatted_total = format_size(total_size)
        st.markdown(
            f'<div style="margin-top: 1rem; margin-bottom: 2rem;">'
            f'<span class="stat-badge">Total: {len(uploaded_files)} file(s)</span>'
            f'<span class="stat-badge">Size: {formatted_total}</span>'
            f'</div>',
            unsafe_allow_html=True,
        )

    # Put an image on the left side to balance empty space
    st.markdown("<br>", unsafe_allow_html=True)
    import base64
    img_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "banner.png")
    with open(img_path, "rb") as img_file:
        img_b64 = base64.b64encode(img_file.read()).decode()
    st.markdown(
        f'<div style="text-align:center;">'
        f'<img src="data:image/png;base64,{img_b64}" style="width:100%; border-radius:12px;" />'
        f'<p style="color:#94a3b8; font-size:0.85rem; margin-top:0.5rem;">Fast & Secure Conversion</p>'
        f'</div>',
        unsafe_allow_html=True,
    )

with right_col:
    st.subheader("Preview & Convert")
    
    if not uploaded_files:
        st.markdown("""
        <div style="background: rgba(30, 41, 59, 0.6); padding: 1.5rem; border-radius: 12px; border: 1px solid rgba(102, 126, 234, 0.3); margin-top: 0.5rem;">
            <h4 style="margin-top:0; color: #f8fafc; font-weight: 600; margin-bottom: 1rem;">How it works</h4>
            <div style="display: flex; align-items: center; margin-bottom: 1rem;">
                <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; width: 26px; height: 26px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 0.85rem; font-weight: bold; margin-right: 12px; flex-shrink: 0;">1</div>
                <span style="color:#cbd5e1; font-size: 0.95rem;"><strong>Upload:</strong> Drag & drop your .docx files on the left.</span>
            </div>
            <div style="display: flex; align-items: center; margin-bottom: 1rem;">
                <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; width: 26px; height: 26px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 0.85rem; font-weight: bold; margin-right: 12px; flex-shrink: 0;">2</div>
                <span style="color:#cbd5e1; font-size: 0.95rem;"><strong>Preview:</strong> Check the loaded files and their total size.</span>
            </div>
            <div style="display: flex; align-items: center;">
                <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; width: 26px; height: 26px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 0.85rem; font-weight: bold; margin-right: 12px; flex-shrink: 0;">3</div>
                <span style="color:#cbd5e1; font-size: 0.95rem;"><strong>Convert:</strong> Click 'Convert' and download them instantly!</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        with st.expander("Uploaded Files", expanded=True):
            for f in uploaded_files:
                formatted_f_size = format_size(f.size)
                st.markdown(
                    f'<div class="file-card"><strong>{f.name}</strong> '
                    f'<span style="color:#888;">({formatted_f_size})</span></div>',
                    unsafe_allow_html=True,
                )
                
        st.markdown("<br>", unsafe_allow_html=True)

        # ─── Convert Button ──────────────────────────────────────────
        if st.button("Convert All to PDF", use_container_width=True):
            progress_bar = st.progress(0)
            status_text = st.empty()

            pdf_files = convert_word_to_pdf(uploaded_files, progress_bar, status_text)

            if pdf_files:
                st.session_state["pdf_files"] = pdf_files

        # ─── Download Section ────────────────────────────────────────
        if "pdf_files" in st.session_state and st.session_state["pdf_files"]:
            pdf_files = st.session_state["pdf_files"]

            st.markdown(
                '<div class="success-box">'
                f"<h3>{len(pdf_files)} PDF(s) Ready!</h3>"
                "<p>Download individually or grab them all as a ZIP.</p>"
                "</div>",
                unsafe_allow_html=True,
            )

            # Individual downloads (2 columns within the right column)
            dl_cols = st.columns(2)
            for i, (name, data) in enumerate(pdf_files):
                with dl_cols[i % 2]:
                    st.download_button(
                        label=f"Download {name}",
                        data=data,
                        file_name=name,
                        mime="application/pdf",
                        key=f"dl_{i}",
                    )

            # ZIP download
            st.markdown("---")
            zip_buffer = create_zip(pdf_files)
            st.download_button(
                label="Download All as ZIP",
                data=zip_buffer,
                file_name="converted_pdfs.zip",
                mime="application/zip",
                key="dl_zip",
                use_container_width=True,
            )

# ─── Features Footer ─────────────────────────────────────────────────
st.markdown("<br><hr style='border: none; border-top: 1px solid rgba(255,255,255,0.1); margin-top: 3rem;'>", unsafe_allow_html=True)

feat1, feat2, feat3 = st.columns(3)
with feat1:
    st.markdown("<div style='text-align: center; color: #94a3b8;'><h4 style='color: #f8fafc;'>Lightning Fast</h4><p style='font-size: 0.85rem;'>Converts bulk files in seconds</p></div>", unsafe_allow_html=True)
with feat2:
    st.markdown("<div style='text-align: center; color: #94a3b8;'><h4 style='color: #f8fafc;'>100% Private</h4><p style='font-size: 0.85rem;'>Processes entirely on your local machine</p></div>", unsafe_allow_html=True)
with feat3:
    st.markdown("<div style='text-align: center; color: #94a3b8;'><h4 style='color: #f8fafc;'>Batch Downloads</h4><p style='font-size: 0.85rem;'>Get all your PDFs neatly zipped</p></div>", unsafe_allow_html=True)
