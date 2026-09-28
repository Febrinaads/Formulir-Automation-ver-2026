"""
app.py  —  Form Penjelasan Dokumentasi → PowerPoint (Streamlit)

Jalankan:
    pip install -r requirements.txt
    streamlit run app.py
"""
from datetime import time

import streamlit as st

from logos import LOGO_KANAN, LOGO_KIRI
from pptx_builder import FormData, ImageItem, build_pptx, count_words, strip_markup

MAX_URAIAN = 250
MAX_CAPTION = 20
MAX_IMAGES = 3
DEFAULT_JENIS = (
    "PCB – (Learning System) Sosialisasi dan Partisipasi – "
    "Bukti Pemahaman Melalui Quiz ONE Pertamina dan Program Budaya"
)

st.set_page_config(page_title="Form Penjelasan Dokumentasi", page_icon="📝", layout="centered")


# ----------------------------------------------------------------------------
# Helper
# ----------------------------------------------------------------------------
def counter(n: int, limit: int, key_suffix: str = ""):
    """Tampilkan penghitung kata (merah jika melebihi batas)."""
    if n > limit:
        st.markdown(f":red[**{n} / {limit} kata — melebihi batas**]")
    else:
        st.caption(f"{n} / {limit} kata")


# ----------------------------------------------------------------------------
# Form
# ----------------------------------------------------------------------------
st.title("Form Penjelasan Dokumentasi")
st.write("Isi formulir, lalu klik **Generate PowerPoint**.")

st.subheader("Jenis Kegiatan")
jenis = st.text_input("Jenis Kegiatan *", value=DEFAULT_JENIS)

st.subheader("Data Narasumber")
c1, c2 = st.columns(2)
nama = c1.text_input("Nama Lengkap *")
nik = c2.text_input("NIK *")
jabatan = c1.text_input("Jabatan *")
email = c2.text_input("Email *")
hsh = c1.text_input("HSH *")
perusahaan = c2.text_input("Perusahaan *")
direktorat = st.text_input("Direktorat *")

st.subheader("Jadwal Pelaksanaan")
tanggal = st.text_input("Hari/Tanggal *", placeholder="contoh: 17-18 May 2026 / Senin, 17 Mei 2026")
t1, t2 = st.columns(2)
waktu_mulai = t1.time_input("Waktu mulai (WIB) *", value=time(9, 0), step=300)
waktu_selesai = t2.time_input("Waktu selesai (WIB) *", value=time(10, 0), step=300)

st.subheader("Informasi")
topik = st.text_input("Topik / Judul Materi *", max_chars=220)
uraian = st.text_area(
    "Uraian Singkat (maks. 250 kata) *",
    height=260,
    help="Bungkus teks dengan **dua bintang** agar tampil merah di slide. Enter = paragraf baru.",
)
n_uraian = count_words(strip_markup(uraian))
counter(n_uraian, MAX_URAIAN)
st.caption("Tips: `**teks**` akan tampil berwarna merah di slide.")

st.subheader("Foto / Gambar (tampil di sisi kanan slide)")
files = st.file_uploader(
    f"Unggah gambar (maks. {MAX_IMAGES})",
    type=["png", "jpg", "jpeg", "webp"],
    accept_multiple_files=True,
)
if files and len(files) > MAX_IMAGES:
    st.warning(f"Hanya {MAX_IMAGES} gambar pertama yang dipakai.")
    files = files[:MAX_IMAGES]

captions = []
caption_over = False
for i, f in enumerate(files or []):
    col_img, col_cap = st.columns([1, 2])
    col_img.image(f, use_container_width=True)
    cap = col_cap.text_input(f"Deskripsi gambar {i + 1} (maks. 20 kata)", key=f"cap_{i}_{f.name}")
    n = count_words(cap)
    with col_cap:
        counter(n, MAX_CAPTION)
    caption_over = caption_over or n > MAX_CAPTION
    captions.append(cap)

st.divider()

# ----------------------------------------------------------------------------
# Generate
# ----------------------------------------------------------------------------
if st.button("Generate PowerPoint", type="primary"):
    st.session_state.pop("pptx_bytes", None)

    required = {
        "Jenis Kegiatan": jenis, "Nama Lengkap": nama, "NIK": nik, "Jabatan": jabatan,
        "Email": email, "HSH": hsh, "Perusahaan": perusahaan, "Direktorat": direktorat,
        "Hari/Tanggal": tanggal, "Topik / Judul Materi": topik, "Uraian Singkat": uraian,
    }
    missing = [k for k, v in required.items() if not v.strip()]

    errors = []
    if missing:
        errors.append("Kolom wajib belum diisi: " + ", ".join(missing))
    if email.strip() and ("@" not in email or "." not in email.split("@")[-1]):
        errors.append("Format email tidak valid.")
    if n_uraian > MAX_URAIAN:
        errors.append(f"Uraian Singkat melebihi {MAX_URAIAN} kata ({n_uraian} kata).")
    if not files:
        errors.append("Unggah minimal satu gambar.")
    if caption_over:
        errors.append(f"Ada deskripsi gambar yang melebihi {MAX_CAPTION} kata.")

    if errors:
        for e in errors:
            st.error(e)
    else:
        data = FormData(
            jenis_kegiatan=jenis.strip(),
            nama=nama.strip(),
            nik=nik.strip(),
            jabatan=jabatan.strip(),
            email=email.strip(),
            hsh=hsh.strip(),
            perusahaan=perusahaan.strip(),
            direktorat=direktorat.strip(),
            tanggal=tanggal.strip(),
            waktu_mulai=waktu_mulai.strftime("%H:%M"),
            waktu_selesai=waktu_selesai.strftime("%H:%M"),
            topik=topik.strip(),
            uraian=uraian.strip(),
            images=[ImageItem(data=f.getvalue(), caption=c.strip()) for f, c in zip(files, captions)],
            logo_kiri=LOGO_KIRI,
            logo_kanan=LOGO_KANAN,
        )
        try:
            with st.spinner("Membuat PowerPoint…"):
                st.session_state["pptx_bytes"] = build_pptx(data)
            safe = "".join(ch for ch in nama if ch not in '\\/:*?"<>|').strip() or "Narasumber"
            st.session_state["pptx_name"] = f"Form Penjelasan Dokumentasi - {safe}.pptx"
        except Exception as exc:  # noqa: BLE001
            st.error(f"Gagal membuat PowerPoint: {exc}")

if "pptx_bytes" in st.session_state:
    st.success("PowerPoint berhasil dibuat.")
    st.download_button(
        "⬇️ Download PowerPoint (.pptx)",
        data=st.session_state["pptx_bytes"],
        file_name=st.session_state["pptx_name"],
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
