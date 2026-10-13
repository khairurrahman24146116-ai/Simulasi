import io

import pandas as pd
import streamlit as st

import analisa as an
import tahap1

st.set_page_config(page_title="Analisis Terminal Kontainer", layout="wide")
st.title("📦 Analisis Event Log Terminal Kontainer")
st.caption("Simulasi & Pemodelan · Tahap 1 (pengambilan nilai)")


@st.cache_data(show_spinner="Membaca data...")
def muat(isi: bytes, nama: str):
    return an.baca_data(isi, nama)


@st.cache_data(show_spinner="Menghitung jalur dan trace...")
def proses(df):
    return an.bangun_trace(df)


def ke_excel(**sheets) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nama, d in sheets.items():
            d.to_excel(w, sheet_name=nama[:31], index=False)
    return buf.getvalue()


file = st.sidebar.file_uploader("Upload data (CSV disarankan, .xlsx juga bisa)",
                                type=["csv", "xlsx"])
st.sidebar.markdown("Kolom wajib: **Case ID**, **Activity**, **Time**")

if file is None:
    st.info("Upload file data di panel kiri untuk mulai. "
            "Tips: simpan Excel sebagai CSV dulu agar jauh lebih cepat dibaca.")
    st.stop()

try:
    df, info = muat(file.getvalue(), file.name)
except ValueError as e:
    st.error(str(e))
    st.stop()

kasus, trace = proses(df)

tab0, tab1, tab3, tab5, tab6, tab_tunggu = st.tabs(
    ["Ringkasan", "Langkah 1 & 4: Jalur / Trace", "Langkah 3: Aktivitas",
     "Langkah 5: Trace per periode", "Langkah 6: Trace 2, 8, 12",
     "Menunggu kepastian"])

with tab0:
    c = st.columns(4)
    c[0].metric("Baris data dipakai", f"{info['baris_dipakai']:,}")
    c[1].metric("Kontainer (Case ID)", f"{len(kasus):,}")
    c[2].metric("Jumlah jalur / trace", f"{len(trace):,}")
    c[3].metric("Aktivitas unik", f"{df['Activity'].nunique():,}")
    st.write(f"Rentang waktu data: **{df['Time'].min():%d/%m/%Y %H:%M}** "
             f"s/d **{df['Time'].max():%d/%m/%Y %H:%M}**")
    if info["waktu_invalid"] or info["baris_kosong"]:
        st.warning(f"Baris dibuang: {info['waktu_invalid']:,} dengan Time tidak terbaca, "
                   f"{info['baris_kosong']:,} dengan Case ID/Activity kosong "
                   f"(dari {info['baris_awal']:,} baris awal).")
    else:
        st.success("Semua baris terbaca dengan benar.")
    st.caption("Cek: jumlah kolom 'Jumlah Kontainer' di tabel trace harus sama dengan "
               f"total kontainer ({len(kasus):,}) → "
               f"{'✅ sama' if trace['Jumlah Kontainer'].sum() == len(kasus) else '❌ BEDA'}")

with tab1:
    st.subheader("Daftar trace (urutan aktivitas per kontainer)")
    st.dataframe(trace, use_container_width=True, hide_index=True, height=350)
    st.download_button("⬇️ Download tabel trace (Excel)",
                       ke_excel(Trace=trace), "hasil_trace.xlsx")
    n = st.slider("Tampilkan N trace teratas di grafik", 3, min(30, len(trace)),
                  min(10, len(trace)))
    st.bar_chart(trace.head(n).set_index("Trace")["Jumlah Kontainer"])

    st.subheader("Lihat detail satu trace")
    pilih = st.selectbox("Pilih trace", trace["Trace"])
    baris = trace[trace["Trace"] == pilih].iloc[0]
    st.write(f"**{pilih}**: {baris['Jumlah Kontainer']:,} kontainer "
             f"({baris['Persen (%)']}%), {baris['Jumlah Aktivitas']} aktivitas")
    langkah = pd.DataFrame({"Urutan": range(1, baris["Jumlah Aktivitas"] + 1),
                            "Aktivitas": baris["Jalur"].split(an.PANAH)})
    a, b = st.columns(2)
    a.dataframe(langkah, hide_index=True, use_container_width=True)
    ids = kasus.loc[kasus["Trace"] == pilih, ["Case ID", "Mulai", "Selesai"]]
    b.dataframe(ids, hide_index=True, use_container_width=True)

with tab3:
    st.subheader("Aktivitas unik dan frekuensinya")
    akt = an.tabel_aktivitas(df)
    st.dataframe(akt, use_container_width=True, hide_index=True)
    st.download_button("⬇️ Download tabel aktivitas (Excel)",
                       ke_excel(Aktivitas=akt), "hasil_aktivitas.xlsx")

with tab5:
    st.subheader("Jumlah kontainer per trace dalam blok waktu")
    c1, c2 = st.columns(2)
    bulan = c1.number_input("Lebar blok (bulan)", 1, 12, 3)
    basis = c2.radio("Kasus dihitung berdasarkan waktu", ["Mulai", "Selesai"],
                     horizontal=True)
    per = an.trace_per_periode(kasus, int(bulan), basis)
    st.dataframe(per, use_container_width=True, hide_index=True)
    st.download_button("⬇️ Download tabel per periode (Excel)",
                       ke_excel(PerPeriode=per), "hasil_trace_per_periode.xlsx")
    kolom_p = [k for k in per.columns if k.startswith("P")]
    st.line_chart(per.set_index("Trace")[kolom_p].head(10).T)
    st.caption("Grafik: 10 trace teratas. Blok dihitung dari bulan pertama data, "
               "bukan dari awal tahun kalender.")

with tab6:
    tahap1.tampilkan(st, file.getvalue(), file.name, df, kasus)

with tab_tunggu:
    st.markdown("""
Bagian ini belum dibuat karena maksud dosen belum jelas:

- **Langkah 2: barang dalam kontainer**: kemungkinan dari kolom `Detail Lampiran`
  (sekarang dipakai di tab Langkah 6 untuk menentukan trace dosen).
- **Langkah 7: analisis setiap trace**: belum ada arahan dosen.

Sudah jelas: **Langkah 6 "2, 8, 12"** = trace 2, 8, dan 12 di file `AktivitasPerTrace`
(lihat tab Langkah 6).
""")