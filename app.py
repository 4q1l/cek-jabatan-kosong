import streamlit as st
import pandas as pd
import io
import asyncio
import zipfile
from playwright.async_api import async_playwright
import os
import sys
import subprocess

# --- 1. SET PAGE CONFIG (WAJIB DI BARIS PERTAMA KODE STREAMLIT & HANYA 1 KALI) ---
st.set_page_config(layout="wide", page_title="Monitoring Jabatan ASN")

# --- 2. CONFIGURASI PLAYWRIGHT ENVIRONMENT ---
PLAYWRIGHT_DIR = os.path.join(os.getcwd(), ".playwright-browsers")
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = PLAYWRIGHT_DIR

@st.cache_resource
def install_playwright_browsers():
    if not os.path.exists(PLAYWRIGHT_DIR) or len(os.listdir(PLAYWRIGHT_DIR)) == 0:
        try:
            subprocess.run(
                [sys.executable, "-m", "playwright", "install", "chromium"],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE
            )
        except Exception as e:
            st.error(f"Gagal mengunduh browser engine: {e}")

# Jalankan instalasi di awal
install_playwright_browsers()

# --- 3. CSS KHUSUS TAMPILAN WEB STREAMLIT ---
st.markdown("""
    <style>
    .block-container { padding-top: 2rem; }
    
    /* Gaya Judul Bagan */
    .judul-bagan {
        text-align: center;
        font-family: 'Segoe UI', Arial, sans-serif;
        margin-bottom: 30px;
        line-height: 1.3;
    }
    .judul-main { font-size: 22px; font-weight: 800; color: #1a252f; letter-spacing: 1px; margin: 0; }
    .judul-sub { font-size: 20px; font-weight: 700; color: #34495e; margin: 5px 0 0 0; text-transform: uppercase; }
    
    /* Tata Letak Pohon Organisasi */
    .tree-container { display: flex; flex-direction: column; align-items: center; justify-content: center; min-width: max-content; }
    .tree ul { padding-top: 40px; position: relative; display: flex; justify-content: center; vertical-align: top; }
    
    /* Setiap rumpun struktur (li) */
    .tree li { float: left; text-align: center; list-style-type: none; position: relative; padding: 40px 5px 0 5px; vertical-align: top; }
    
    /* Garis Horizontal Penghubung Antar Kotak */
    .tree li::before, .tree li::after { content: ''; position: absolute; top: 0; right: 50%; border-top: 2px solid #ccc; width: 50%; height: 40px; }
    .tree li::after { right: auto; left: 50%; border-left: 2px solid #ccc; }
    
    /* Menghilangkan garis pembantu untuk kotak tunggal */
    .tree li:only-child::after, .tree li:only-child::before { display: none; }
    .tree li:only-child { padding-top: 0; }
    .tree li:first-child::before, .tree li:last-child::after { border: 0 none; }
    .tree li:last-child::before { border-right: 2px solid #ccc; border-radius: 0 5px 0 0; }
    .tree li:first-child::after { border-radius: 5px 0 0 0; }
    
    /* Garis Vertikal Utama dari Atasan ke Cabang Bawahan */
    .tree ul ul::before { content: ''; position: absolute; top: 0; left: 50%; border-left: 2px solid #ccc; width: 0; height: 40px; }
    
    /* Node Card / Box */
    .node-card { 
        border: 1px solid #ccc; 
        padding: 12px 10px; 
        display: inline-block; 
        border-radius: 8px; 
        width: 200px; 
        max-width: 220px; 
        word-wrap: break-word; 
        white-space: normal; 
        margin: 0px 10px 10px 10px; 
        background-color: white; 
        box-shadow: 0 2px 4px rgba(0,0,0,0.1); 
        text-align: center; 
        position: relative; 
        z-index: 10; 
    }
    .terisi { border-top: 8px solid #2ecc71; }
    .kosong { border-top: 8px solid #e74c3c; background-color: #fff5f5; }
    
    .text-jabatan { 
        font-size: 12px; 
        font-weight: bold; 
        color: #2c3e50; 
        margin-bottom: 6px; 
        line-height: 1.4;
        white-space: normal !important; 
    }
    .text-nama { font-size: 11px; color: #34495e; margin: 0; line-height: 1.3; white-space: normal !important; }
    .text-nip { font-size: 10px; color: #7f8c8d; margin: 2px 0 0 0; }
    .text-pangkat { font-size: 10px; color: #2c3e50; margin: 1px 0 4px 0; font-style: italic; }
    .text-eselon { font-size: 9px; background: #ebf2ff; padding: 2px 5px; border-radius: 4px; color: #2980b9; font-weight: bold; display: inline-block; margin-top: 5px;}
    
    .li-level-a { padding-top: 40px !important; }
    .li-level-b { padding-top: 100px !important; }
    .li-level-b::after { height: 100px !important; }
    .li-level-b::before { height: 100px !important; }
    
    .scroll-web { overflow-x: auto; padding: 20px; background: #ffffff; border: 1px dashed #ddd; border-radius: 8px; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 Monitoring Struktur Organisasi")

# --- SIDEBAR KONFIGURASI KOLOM EXCEL ---
st.sidebar.header("Konfigurasi Kolom Excel")
col_unit_master = st.sidebar.text_input("Kolom Unit Kerja di Master", "Unit Kerja")
col_jab_master = st.sidebar.text_input("Kolom Jabatan di Master", "Nama Jabatan")
col_atasan_master = st.sidebar.text_input("Kolom Atasan di Master", "Atasan")
col_eselon_master = st.sidebar.text_input("Kolom Eselon di Master", "Eselon")

st.sidebar.markdown("---")
col_jab_asn = st.sidebar.text_input("Kolom Jabatan di Real ASN", "jabatannama")
col_nama = st.sidebar.text_input("Kolom Nama Pegawai", "namalengkap")
col_nip = st.sidebar.text_input("Kolom NIP", "nip")
col_golongan = st.sidebar.text_input("Kolom Golongan Pegawai", "golruangnama")
col_pangkat = st.sidebar.text_input("Kolom Pangkat Pegawai", "pangkat")

st.sidebar.markdown("---")
file_asn = st.sidebar.file_uploader("1. Upload Data Real ASN (A-DT)", type=["xlsx"])
file_master = st.sidebar.file_uploader("2. Upload Peta Jabatan Master Baru", type=["xlsx"])

# --- FUNGSI KLASIFIKASI KETINGGIAN KOTAK (LI) ---
def get_li_eselon_class(eselon_val):
    if pd.isna(eselon_val):
        return "li-level-a"
    eselon_str = str(eselon_val).strip().lower()
    if eselon_str.endswith('b') or 'b' in eselon_str:
        return "li-level-b"
    return "li-level-a"

def get_eselon_weight(eselon_val):
    if pd.isna(eselon_val): return 99
    eselon_str = str(eselon_val).strip().lower()
    if '2a' in eselon_str or 'ii.a' in eselon_str: return 1
    elif '2b' in eselon_str or 'ii.b' in eselon_str: return 2
    elif '3a' in eselon_str or 'iii.a' in eselon_str: return 3
    elif '3b' in eselon_str or 'iii.b' in eselon_str: return 4
    elif '4a' in eselon_str or 'iv.a' in eselon_str: return 5
    elif '4b' in eselon_str or 'iv.b' in eselon_str: return 6
    return 90

# --- RECURSION TREE GENERATOR ---
def build_tree_html(df, parent_name=None):
    if parent_name is None or pd.isna(parent_name) or str(parent_name).strip() == "":
        mask = (
            df[col_atasan_master].isna() | 
            (df[col_atasan_master].astype(str).str.strip() == "") | 
            (df[col_atasan_master].astype(str).str.strip().str.upper() == "NAN")
        )
        children = df[mask]
    else:
        parent_clean = str(parent_name).strip().upper()
        children = df[df[col_atasan_master].astype(str).str.strip().str.upper() == parent_clean]
    
    if children.empty: 
        return ""
    
    children = children.copy()
    children['eselon_weight'] = children[col_eselon_master].apply(get_eselon_weight)
    children = children.sort_values(by='eselon_weight')
    
    html = "<ul>"
    for _, row in children.iterrows():
        is_empty = pd.isna(row[col_nama]) or str(row[col_nama]).strip() == "" or str(row[col_nama]).strip().upper() == "NAN"
        status_class = "kosong" if is_empty else "terisi"
        
        if is_empty:
            nama_display = "❌ KOSONG"
            detail_html = f'<p class="text-nama">{nama_display}</p>'
        else:
            nama_display = row[col_nama]
            nip_display = row[col_nip] if not pd.isna(row[col_nip]) else "-"
            gol_val = row[col_golongan] if col_golongan in row and not pd.isna(row[col_golongan]) else "-"
            pakt_val = row[col_pangkat] if col_pangkat in row and not pd.isna(row[col_pangkat]) else "-"
            eselon_val = row[col_eselon_master] if col_eselon_master in row and not pd.isna(row[col_eselon_master]) else "-"
            
            detail_html = (
                f'<p class="text-nama">{nama_display}</p>'
                f'<p class="text-pangkat">{pakt_val} ({gol_val})</p>'
                f'<p class="text-nip">NIP: {nip_display}</p>'
                f'<span class="text-eselon">{eselon_val}</span>'
            )
            
        eselon_val_raw = row[col_eselon_master] if col_eselon_master in row and not pd.isna(row[col_eselon_master]) else "-"
        li_position_class = get_li_eselon_class(eselon_val_raw)

        html += f'<li class="{li_position_class}">'
        html += (
            f'<div class="node-card {status_class}">'
            f'<div class="text-jabatan">{row[col_jab_master]}</div>'
            f'{detail_html}'
            f'</div>'
        )
        html += build_tree_html(df, row[col_jab_master])
        html += "</li>"
    html += "</ul>"
    return html

# --- HELPER TEMPLATE HTML UNTUK PDF ---
def get_full_html_document(title_unit, tree_content):
    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            body {{ font-family: 'Segoe UI', Arial, sans-serif; background-color: #ffffff; margin: 0; padding: 40px; }}
            .judul-bagan {{ text-align: center; margin-bottom: 35px; line-height: 1.4; }}
            .judul-main {{ font-size: 24px; font-weight: 800; color: #000000; letter-spacing: 1px; margin: 0; }}
            .judul-sub {{ font-size: 21px; font-weight: 700; color: #222222; margin: 6px 0 0 0; text-transform: uppercase; }}
            
            .tree-container {{ display: inline-block; text-align: center; min-width: max-content; margin: 0 auto; }}
            .tree ul {{ padding-top: 40px; position: relative; display: flex; justify-content: center; margin: 0; padding-left: 0; }}
            .tree li {{ float: left; text-align: center; list-style-type: none; position: relative; padding: 40px 5px 0 5px; vertical-align: top; }}
            .tree li::before, .tree li::after{{ content: ''; position: absolute; top: 0; right: 50%; border-top: 2px solid #555555; width: 50%; height: 40px; }}
            .tree li::after{{ right: auto; left: 50%; border-left: 2px solid #555555; }}
            .tree li:only-child::after, .tree li:only-child::before {{ display: none; }}
            .tree li:only-child{{ padding-top: 0; }}
            .tree li:first-child::before, .tree li:last-child::after{{ border: 0 none; }}
            .tree li:last-child::before{{ border-right: 2px solid #555555; border-radius: 0 5px 0 0; }}
            .tree li:first-child::after{{ border-radius: 5px 0 0 0; }}
            .tree ul ul::before{{ content: ''; position: absolute; top: 0; left: 50%; border-left: 2px solid #555555; width: 0; height: 40px; }}
            
            .node-card {{ 
                border: 1px solid #444444; 
                padding: 12px 10px; 
                display: inline-block; 
                border-radius: 8px; 
                width: 190px; 
                max-width: 200px;
                white-space: normal !important;
                word-wrap: break-word;
                background-color: #ffffff !important; 
                box-shadow: 0 2px 4px rgba(0,0,0,0.1); 
                text-align: center; 
                position: relative; 
                z-index: 10; 
                margin: 0 10px 10px 10px; 
            }}
            .terisi {{ border-top: 8px solid #2ecc71 !important; }}
            .kosong {{ border-top: 8px solid #e74c3c !important; background-color: #fff5f5 !important; }}
            .text-jabatan {{ font-size: 12px; font-weight: bold; color: #000000 !important; margin: 0 0 6px 0; line-height: 1.4; white-space: normal !important; }}
            .text-nama {{ font-size: 11px; color: #222222 !important; margin: 0 0 4px 0; font-weight: 500; white-space: normal !important; line-height: 1.3; }}
            .text-pangkat {{ font-size: 10px; color: #111111 !important; margin: 1px 0 5px 0; font-style: italic; white-space: normal !important; }}
            .text-nip {{ font-size: 10px; color: #555555 !important; margin: 0 0 6px 0; }}
            .text-eselon {{ font-size: 9px; background-color: #ebf2ff !important; padding: 2px 6px; border-radius: 4px; color: #1a56db !important; font-weight: bold; display: inline-block; border: 1px solid #b3d1ff; }}
            .tree {{ display: block; width: max-content; margin: 0 auto; }}
            
            .li-level-a {{ padding-top: 40px !important; }}
            .li-level-b {{ padding-top: 100px !important; }}
            .li-level-b::after {{ height: 100px !important; }}
            .li-level-b::before {{ height: 100px !important; }}
        </style>
    </head>
    <body>
        <div class="tree-container">
            <div class="judul-bagan">
                <div class="judul-main">STRUKTUR ORGANISASI</div>
                <div class="judul-sub">{title_unit}</div>
            </div>
            <div class="tree">
                {tree_content}
            </div>
        </div>
    </body>
    </html>
    """

# --- BACKEND FUNCTION: GENERATOR SINGLE PDF ---
async def generate_pdf_from_html(html_content):
    install_playwright_browsers()
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=[
                '--no-sandbox', 
                '--disable-setuid-sandbox', 
                '--disable-dev-shm-usage',
                '--disable-gpu'
            ]
        )
        page = await browser.new_page(viewport={"width": 7000, "height": 3000})
        await page.set_content(html_content, wait_until="networkidle")
        
        dimensions = await page.evaluate("""() => {
            const el = document.querySelector('.tree-container');
            return {
                width: el ? el.getBoundingClientRect().width + 200 : 2500,
                height: el ? el.getBoundingClientRect().height + 300 : 1800
            }
        }""")
        
        width_in_inches = dimensions['width'] / 96
        height_in_inches = dimensions['height'] / 96
        
        pdf_bytes = await page.pdf(
            print_background=True,
            width=f"{width_in_inches}in",
            height=f"{height_in_inches}in",
            margin={"top": "0.6in", "right": "0.6in", "bottom": "0.6in", "left": "0.6in"}
        )
        await browser.close()
        return pdf_bytes

# --- BACKEND FUNCTION: BATCH GENERATOR ZIP ---
async def generate_all_charts_zip(df_master_merge, col_unit, build_tree_fn):
    install_playwright_browsers()
    zip_buffer = io.BytesIO()
    unit_list = df_master_merge[col_unit].dropna().unique()
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--disable-gpu']
        )
        page = await browser.new_page(viewport={"width": 7000, "height": 3000})
        
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for index, unit in enumerate(unit_list):
                status_text.text(f"Memproses ({index+1}/{len(unit_list)}): {unit}")
                df_unit = df_master_merge[df_master_merge[col_unit] == unit]
                tree_html = build_tree_fn(df_unit)
                
                full_html = get_full_html_document(unit, tree_html)
                await page.set_content(full_html, wait_until="networkidle")
                
                dimensions = await page.evaluate("""() => {
                    const el = document.querySelector('.tree-container');
                    return { 
                        width: el ? el.getBoundingClientRect().width + 200 : 2500, 
                        height: el ? el.getBoundingClientRect().height + 300 : 1800 
                    }
                }""")
                
                pdf_bytes = await page.pdf(
                    print_background=True,
                    width=f"{dimensions['width'] / 96}in",
                    height=f"{dimensions['height'] / 96}in",
                    margin={"top": "0.6in", "right": "0.6in", "bottom": "0.6in", "left": "0.6in"}
                )
                
                clean_filename = f"Bagan_Struktur_{str(unit).replace(' ', '_').replace('/', '-')}.pdf"
                zip_file.writestr(clean_filename, pdf_bytes)
                progress_bar.progress((index + 1) / len(unit_list))
            
            status_text.text("✨ Sukses! Seluruh file bagan berhasil dikompresi.")
            progress_bar.empty()
        await browser.close()
        
    return zip_buffer.getvalue()

# Eksekusi Pemrosesan File
if file_asn and file_master:
    df_asn = pd.read_excel(file_asn)
    df_master = pd.read_excel(file_master)

    # Pembersihan karakter _x000D_ dan Enter tersembunyi
    for col in [col_unit_master, col_jab_master, col_atasan_master, col_eselon_master]:
        if col in df_master.columns:
            df_master[col] = df_master[col].astype(str).str.replace(r'[\r\n]', '', regex=True)
            df_master[col] = df_master[col].astype(str).str.replace('_x000D_', '', regex=True, case=False).str.strip()

    cols_to_clean_asn = [col_jab_asn, col_nama, col_nip, col_golongan, col_pangkat]
    for col in cols_to_clean_asn:
        if col in df_asn.columns:
            df_asn[col] = df_asn[col].astype(str).str.replace(r'[\r\n]', '', regex=True)
            df_asn[col] = df_asn[col].astype(str).str.replace('_x000D_', '', regex=True, case=False).str.strip()

    df_asn[col_jab_asn] = df_asn[col_jab_asn].astype(str).str.upper()
    df_master[col_jab_master] = df_master[col_jab_master].astype(str).str.upper()
    df_master[col_atasan_master] = df_master[col_atasan_master].astype(str).str.upper()

    df_merge = pd.merge(df_master, df_asn[[col_jab_asn, col_nama, col_nip, col_golongan, col_pangkat]], 
                        left_on=col_jab_master, right_on=col_jab_asn, how='left')
    
    df_merge['Status'] = df_merge[col_nama].apply(lambda x: "TERISI" if not pd.isna(x) and str(x).strip() != "" and str(x).strip().upper() != "NAN" else "KOSONG")

    tab1, tab2 = st.tabs(["Visualisasi Bagan bertingkat", "📋 Rekapitulasi & Download"])

    with tab1:
        if col_unit_master in df_merge.columns:
            unit_list = df_merge[col_unit_master].unique()
            
            st.markdown("### 📦 Ekspor Masal Seluruh Struktur Dinas")
            st.write("Klik tombol di bawah untuk membuat dan mengunduh bagan PDF seluruh OPD sekaligus dalam satu file kompresi.")
            
            if st.button("📥 Download Semua Bagan (Format ZIP/RAR)", type="secondary", key="download_all_zip"):
                with st.spinner("Sistem sedang merender seluruh bagan dinas secara latar belakang... Harap tunggu sebentar..."):
                    zip_data = asyncio.run(generate_all_charts_zip(df_merge, col_unit_master, build_tree_html))
                    st.download_button(
                        label="💾 Klik Di Sini Untuk Menyimpan File ZIP",
                        data=zip_data,
                        file_name="Seluruh_Bagan_Struktur_OPD.zip",
                        mime="application/zip"
                    )
            st.markdown("---")
            
            selected_unit = st.selectbox("Pilih Unit Kerja Tampilan:", unit_list)
            df_unit = df_merge[df_merge[col_unit_master] == selected_unit]
            
            tree_html = build_tree_html(df_unit)
            
            st.markdown(f"""
                <div class="scroll-web">
                    <div class="tree-container">
                        <div class="judul-bagan">
                            <div class="judul-main">STRUKTUR ORGANISASI</div>
                            <div class="judul-sub">{selected_unit}</div>
                        </div>
                        <div class="tree">
                            {tree_html}
                        </div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
            
            st.write("") 
            
            html_document_for_pdf = get_full_html_document(selected_unit, tree_html)
            if st.button("📥 Download Bagan Unit Ini Saja (PDF)", type="primary"):
                with st.spinner("Memproses layout PDF..."):
                    pdf_data = asyncio.run(generate_pdf_from_html(html_document_for_pdf))
                    st.download_button(
                        label="Klik di Sini untuk Mengunduh Berkas PDF",
                        data=pdf_data,
                        file_name=f"Bagan_Struktur_{selected_unit.replace(' ', '_')}.pdf",
                        mime="application/pdf"
                    )
        else:
            st.error(f"Kolom '{col_unit_master}' tidak ditemukan di file Master.")

    with tab2:
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Jabatan", len(df_merge))
        m2.metric("Terisi", len(df_merge[df_merge['Status'] == "TERISI"]))
        m3.metric("Kosong", len(df_merge[df_merge['Status'] == "KOSONG"]), delta_color="inverse")

        opsi = st.radio("Tampilkan Data:", ["Semua", "Hanya Kosong", "Hanya Terisi"], horizontal=True)
        df_table = df_merge[df_merge['Status'] == "KOSONG"] if opsi == "Hanya Kosong" else (df_merge[df_merge['Status'] == "TERISI"] if opsi == "Hanya Terisi" else df_merge)
        
        st.dataframe(df_table[[col_unit_master, col_jab_master, col_eselon_master, 'Status', col_nama, col_pangkat, col_golongan]], use_container_width=True)

        def download_excel(df_all):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_all.to_excel(writer, index=False, sheet_name='Semua Jabatan')
                df_all[df_all['Status'] == "KOSONG"].to_excel(writer, index=False, sheet_name='Hanya Kosong')
                df_all[df_all['Status'] == "TERISI"].to_excel(writer, index=False, sheet_name='Hanya Terisi')
            return output.getvalue()

        st.download_button(
            label="📥 Download Rekap Excel (3 Sheet)",
            data=download_excel(df_merge),
            file_name='rekap_jabatan_asn.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
else:
    st.info("Silakan upload kedua file Excel di sidebar untuk memulai.")
