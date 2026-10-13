"""Tahap 1 - hitung kontainer per trace DOSEN (trace 1-12 di file AktivitasPerTrace)
per bulan.

Beda dengan tab "Jalur / Trace" (135 jalur hasil urutan aktivitas), di sini trace
diambil dari kolom `Detail Lampiran` di data, sama seperti cara laporan contoh dosen:
  1. isi sel Detail Lampiran yang kosong,
  2. cocokkan dengan nama trace di AktivitasPerTrace,
  3. hitung aktivitas Document_Entry_via_PDE (1 baris = 1 kontainer) per bulan.
"""
import csv
import io
import re

import pandas as pd

AKTIVITAS_MASUK = "Document_Entry_via_PDE"

# Nama trace persis seperti di file AktivitasPerTrace.xls (sheet trace1..trace12)
TRACE_DOSEN = {
    1: "Quarantine;Dry;Green Line",
    2: "Quarantine;Dry;Red Line",
    3: "Quarantine;Reefer;Green Line",
    4: "Quarantine;Reefer;Red Line",
    5: "Quarantine;Uncontainer;Green Line",
    6: "Quarantine;Uncontainer;Red Line",
    7: "Dry;Green Line",
    8: "Dry;Red Line",
    9: "Reefer;Green Line",
    10: "Reefer;Red Line",
    11: "Uncontainer;Green Line",
    12: "Uncontainer;Red Line",
}
TRACE_KELOMPOK = [2, 8, 12]

BULAN = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
         "Agustus", "September", "Oktober", "November", "Desember"]


def normalisasi(teks) -> str:
    """'Uncontainer; Red Line ' dan 'uncontainer;Red  Line' dianggap sama."""
    t = re.sub(r"\s+", " ", str(teks)).strip()
    t = re.sub(r"\s*;\s*", ";", t)
    return t.casefold()


_KUNCI_DOSEN = {normalisasi(v): k for k, v in TRACE_DOSEN.items()}


def _rapikan_id(seri: pd.Series) -> pd.Series:
    s = seri.astype(str).str.strip()
    return s.str.replace(r"\.0$", "", regex=True)


def _nama_kolom(c) -> str:
    return re.sub(r"[\s_]+", " ", str(c)).strip().casefold()


def _kolom_dipakai(c) -> bool:
    n = _nama_kolom(c)
    return n == "case id" or n == "caseid" or "lampiran" in n


def _baca_mentah(isi: bytes, nama: str) -> pd.DataFrame:
    """Baca hanya kolom Case ID dan Detail Lampiran dari file asli (csv/xlsx)."""
    if nama.lower().endswith(".csv"):
        kepala = isi[:20000]
        for enc in ("utf-8-sig", "latin-1"):
            try:
                teks = kepala.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        baris1 = teks.splitlines()[0] if teks else ""
        try:
            sep = csv.Sniffer().sniff(baris1, delimiters=",;\t|").delimiter
        except csv.Error:
            sep = ","
        return pd.read_csv(io.BytesIO(isi), sep=sep, encoding=enc, dtype=str,
                           usecols=_kolom_dipakai)
    return pd.read_excel(io.BytesIO(isi), dtype=str, usecols=_kolom_dipakai,
                         engine="openpyxl")


def muat_label(isi: bytes, nama: str):
    """Satu label Detail Lampiran per Case ID.

    Return (label, info). label punya kolom: Case ID, Detail Lampiran.
    Label diambil dari nilai terisi pertama di tiap Case ID, jadi tidak
    bergantung pada urutan baris (beda dengan fill-down Excel yang rawan
    kalau data tidak terurut per kontainer).
    """
    d = _baca_mentah(isi, nama)
    kol_id = next((c for c in d.columns if _nama_kolom(c) in ("case id", "caseid")), None)
    kol_lamp = next((c for c in d.columns if "lampiran" in _nama_kolom(c)), None)
    if kol_id is None or kol_lamp is None:
        raise ValueError("Kolom 'Case ID' dan/atau 'Detail Lampiran' tidak ditemukan. "
                         f"Kolom yang terbaca: {list(d.columns)}")
    d = d.rename(columns={kol_id: "Case ID", kol_lamp: "Detail Lampiran"})
    d["Case ID"] = _rapikan_id(d["Case ID"])
    d["Detail Lampiran"] = d["Detail Lampiran"].astype("string").str.strip()
    terisi = d[d["Detail Lampiran"].notna() & (d["Detail Lampiran"] != "")]
    per_kasus = terisi.groupby("Case ID", sort=False)["Detail Lampiran"]
    label = per_kasus.first().reset_index()
    ganda = int((per_kasus.nunique() > 1).sum())
    info = {"kasus_dengan_label": len(label), "kasus_label_ganda": ganda,
            "baris_mentah": len(d)}
    return label, info


def tabel_kasus(df: pd.DataFrame, kasus: pd.DataFrame, label: pd.DataFrame,
                basis: str) -> pd.DataFrame:
    """Satu baris per kontainer: Case ID, Waktu (untuk bulan), Label, Trace dosen."""
    if basis == "entry":
        w = (df.loc[df["Activity"] == AKTIVITAS_MASUK]
             .assign(**{"Case ID": lambda x: _rapikan_id(x["Case ID"])})
             .groupby("Case ID")["Time"].min())
    else:
        k = kasus.assign(**{"Case ID": _rapikan_id(kasus["Case ID"])})
        w = k.set_index("Case ID")["Mulai"]
    semua = pd.DataFrame({"Case ID": _rapikan_id(kasus["Case ID"]).unique()})
    t = (semua.merge(w.rename("Waktu").reset_index(), on="Case ID", how="left")
              .merge(label, on="Case ID", how="left"))
    t["Kunci"] = t["Detail Lampiran"].map(lambda x: normalisasi(x) if pd.notna(x) else None)
    t["Trace dosen"] = t["Kunci"].map(_KUNCI_DOSEN)
    t["Bulan"] = t["Waktu"].dt.to_period("M")
    return t


def _nama_bulan(p: pd.Period) -> str:
    return f"{BULAN[p.month]} {p.year}"


def tabel_hasil(t: pd.DataFrame, pilih: list) -> pd.DataFrame:
    """Tabel akhir: Trace x bulan + Total (format seperti laporan contoh)."""
    bulan = sorted(t["Bulan"].dropna().unique())
    baris = []
    for n in pilih:
        sub = t[t["Trace dosen"] == n]
        r = {"Trace": f"Trace {n}", "Nama trace": TRACE_DOSEN[n]}
        for b in bulan:
            r[_nama_bulan(b)] = int((sub["Bulan"] == b).sum())
        r["Total"] = int(sub["Bulan"].notna().sum())
        baris.append(r)
    out = pd.DataFrame(baris)
    if len(out):
        tot = {"Trace": "TOTAL", "Nama trace": ""}
        for c in out.columns[2:]:
            tot[c] = int(out[c].sum())
        out = pd.concat([out, pd.DataFrame([tot])], ignore_index=True)
    return out


def tabel_label(t: pd.DataFrame) -> pd.DataFrame:
    """Semua nilai Detail Lampiran di data + cocok dengan trace dosen nomor berapa."""
    bulan = sorted(t["Bulan"].dropna().unique())
    pakai = t.assign(Label=t["Detail Lampiran"].fillna("(kosong)"))
    baris = []
    for lab, sub in pakai.groupby("Label", sort=False):
        n = sub["Trace dosen"].dropna()
        r = {"Detail Lampiran": lab,
             "Cocok trace dosen": f"Trace {int(n.iloc[0])}" if len(n) else "-"}
        for b in bulan:
            r[_nama_bulan(b)] = int((sub["Bulan"] == b).sum())
        r["Total kontainer"] = len(sub)
        baris.append(r)
    return (pd.DataFrame(baris).sort_values("Total kontainer", ascending=False)
            .reset_index(drop=True))


def _ke_excel(**sheets) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for nama, d in sheets.items():
            d.to_excel(w, sheet_name=nama[:31], index=False)
    return buf.getvalue()


def tampilkan(st, isi: bytes, nama: str, df: pd.DataFrame, kasus: pd.DataFrame):
    """Isi tab Streamlit. Dipanggil dari app.py."""
    st.subheader("Jumlah kontainer per trace dosen, per bulan")
    st.caption("Trace di sini = 12 trace di file AktivitasPerTrace (dibaca dari kolom "
               "Detail Lampiran), bukan 135 jalur di tab sebelumnya.")

    @st.cache_data(show_spinner="Membaca kolom Detail Lampiran...")
    def _muat(isi_: bytes, nama_: str):
        return muat_label(isi_, nama_)

    try:
        label, info = _muat(isi, nama)
    except ValueError as e:
        st.error(str(e))
        return
    except Exception as e:  # file rusak / format tak terduga
        st.error(f"Gagal membaca Detail Lampiran: {e}")
        return

    c1, c2 = st.columns(2)
    pilih = c1.multiselect(
        "Trace yang dihitung", list(TRACE_DOSEN), default=TRACE_KELOMPOK,
        format_func=lambda n: f"Trace {n}: {TRACE_DOSEN[n]}")
    basis_lbl = c2.radio(
        "Bulan kontainer ditentukan dari",
        ["Waktu Document_Entry_via_PDE (seperti laporan contoh)",
         "Waktu pertama kontainer muncul di data"])
    basis = "entry" if basis_lbl.startswith("Waktu Document") else "mulai"

    t = tabel_kasus(df, kasus, label, basis)
    hasil = tabel_hasil(t, pilih)
    st.dataframe(hasil, use_container_width=True, hide_index=True)
    st.download_button("⬇️ Download tabel trace dosen (Excel)",
                       _ke_excel(TraceDosen=hasil, SemuaLabel=tabel_label(t)),
                       "hasil_trace_dosen.xlsx")

    # --- pengecekan ---
    total = len(t)
    berlabel = int(t["Detail Lampiran"].notna().sum())
    cocok = int(t["Trace dosen"].notna().sum())
    tanpa_waktu = int(t["Bulan"].isna().sum())
    st.caption(f"Cek: {total:,} kontainer di data → {berlabel:,} punya Detail Lampiran, "
               f"{cocok:,} cocok dengan salah satu dari 12 trace dosen.")
    if berlabel < total:
        st.warning(f"{total - berlabel:,} kontainer tidak punya Detail Lampiran sama sekali.")
    if berlabel and cocok < berlabel:
        st.warning(f"{berlabel - cocok:,} kontainer punya Detail Lampiran yang tidak sama "
                   "dengan 12 trace dosen. Lihat tabel di bawah untuk nilai aslinya.")
    if tanpa_waktu:
        st.warning(f"{tanpa_waktu:,} kontainer tidak punya waktu untuk basis yang dipilih "
                   "(tidak masuk hitungan bulan).")
    if info["kasus_label_ganda"]:
        st.warning(f"{info['kasus_label_ganda']:,} kontainer punya lebih dari satu nilai "
                   "Detail Lampiran; dipakai nilai terisi yang pertama.")
    if len(pilih) and hasil.iloc[:-1]["Total"].sum() == 0:
        st.error("Semua trace terpilih bernilai 0: tidak ada Detail Lampiran yang cocok. "
                 "Cek ejaan di tabel bawah.")

    with st.expander("Semua nilai Detail Lampiran di data (untuk cek ejaan)"):
        st.dataframe(tabel_label(t), use_container_width=True, hide_index=True)