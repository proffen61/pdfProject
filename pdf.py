import streamlit as st
from docxtpl import DocxTemplate
from docx import Document
import tempfile
import os
import zipfile

st.set_page_config(page_title="Surat Generator", layout="centered")
st.title("📄 Smart Auto-Fill Surat")

uploaded_template = st.file_uploader("📎 Upload Template Surat (.docx saja)", type=["docx"])

if uploaded_template:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(uploaded_template.read())
        template_path = tmp.name

    with st.form("form_surat"):
        tempatTanggal = st.text_input("🗓️ Tempat & Tanggal (mis: Madiun, 31 Juli 2025)")
        alasan = st.text_area("📌 Alasan")
        peserta = st.text_input("🧑 Nama Peserta (gunakan `;` untuk banyak nama)")
        haritanggal = st.text_input("📆 Hari / Tanggal")
        waktu = st.text_input("⏰ Waktu")
        tempat = st.text_input("📍 Tempat")
        alamat = st.text_input("🏢 Alamat")
        nomor = st.text_input("🆔 Nomor Surat")
        lampiran = st.text_input("📄 Lampiran")
        perihal = st.text_input("✉️ Perihal")
        penandaTangan = st.text_input("✍️ Nama Penandatangan")
        jabatan = st.text_input("🏷️ Jabatan Penandatangan")

        st.caption("ℹ️ Pisahkan beberapa nama peserta dengan tanda titik koma (`;`) untuk menghasilkan banyak surat sekaligus.")
        submitted = st.form_submit_button("🔄 Generate Surat")

    if submitted:
        # Split input into multiple names (if any)
        nama_list = [p.strip() for p in peserta.split(";") if p.strip()]
        first_name = nama_list[0] if nama_list else "[Nama Peserta]"

        # Prepare context for preview (either with first name or a placeholder)
        preview_context = {
            "tempatTanggal": tempatTanggal,
            "alasan": alasan,
            "peserta": first_name,
            "haritanggal": haritanggal,
            "waktu": waktu,
            "tempat": tempat,
            "alamat": alamat,
            "nomor": nomor,
            "lampiran": lampiran,
            "perihal": perihal,
            "penandaTangan": penandaTangan,
            "jabatan": jabatan,
        }

        # Generate preview letter
        doc = DocxTemplate(template_path)
        doc.render(preview_context)
        temp_preview_path = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        doc.save(temp_preview_path.name)

        docx_preview = Document(temp_preview_path.name)
        preview_text = "\n".join([p.text for p in docx_preview.paragraphs if p.text.strip()])
        st.text_area("📄 Preview Surat", value=preview_text or "[Pratinjau kosong: belum ada isi yang dimasukkan.]", height=400)
        st.caption("👀 Ini adalah pratinjau dari surat pertama (atau kosong jika belum diisi).")

        # Generate actual letters
        if len(nama_list) == 1:
            output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
            doc.save(output_path.name)

            st.success("✅ Surat berhasil dibuat!")
            with open(output_path.name, "rb") as f:
                st.download_button(
                    label="⬇️ Download Surat",
                    data=f,
                    file_name=f"Surat_{first_name.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        elif len(nama_list) > 1:
            output_zip = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
            with zipfile.ZipFile(output_zip.name, "w") as zipf:
                for name in nama_list:
                    doc = DocxTemplate(template_path)
                    context = {
                        **preview_context,  # Reuse existing fields
                        "peserta": name     # Override name per letter
                    }
                    doc.render(context)
                    temp_doc = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
                    doc.save(temp_doc.name)
                    filename = f"Surat_{name.replace(' ', '_').replace(',', '')}.docx"
                    zipf.write(temp_doc.name, arcname=filename)

            st.success(f"✅ {len(nama_list)} surat berhasil dibuat.")
            with open(output_zip.name, "rb") as f:
                st.download_button(
                    label="⬇️ Download Semua Surat (ZIP)",
                    data=f,
                    file_name="Surat_Massal.zip",
                    mime="application/zip"
                )
