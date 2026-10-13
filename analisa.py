"""Logika analisis event log terminal kontainer (tanpa tampilan)."""
import io

import pandas as pd

WAJIB = ["Case ID", "Activity", "Time"]
PANAH = " → "


def baca_data(isi: bytes, nama: str):
    """Baca CSV/Excel. Return (DataFrame bersih terurut, info)."""
    nama = nama.lower()
    if nama.endswith((".xlsx", ".xls")):
        df = pd.read_excel(io.BytesIO(isi), usecols=lambda c: str(c).strip() in WAJIB)
    else:
        try:
            teks, enc = isi[:5000].decode("utf-8-sig"), "utf-8-sig"
        except UnicodeDecodeError:
            teks, enc = isi[:5000].decode("latin-1"), "latin-1"
        baris1 = teks.splitlines()[0] if teks else ""
        sep = max([",", ";", "\t"], key=baris1.count)
        df = pd.read_csv(io.BytesIO(isi), sep=sep, encoding=enc,
                         usecols=lambda c: str(c).strip() in WAJIB, dtype=str)

    df.columns = [str(c).strip() for c in df.columns]
    hilang = [c for c in WAJIB if c not in df.columns]
    if hilang:
        raise ValueError(f"Kolom tidak ditemukan: {', '.join(hilang)}. "
                         f"Kolom yang terbaca: {', '.join(df.columns)}")

    df = df[WAJIB].copy()
    baris_awal = len(df)
    df["Case ID"] = df["Case ID"].astype(str).str.strip()
    df["Activity"] = df["Activity"].astype(str).str.strip()
    df["Time"] = pd.to_datetime(df["Time"], dayfirst=True, errors="coerce")
    kosong = df["Case ID"].isin(["", "nan", "None"]) | df["Activity"].isin(["", "nan", "None"])
    waktu_invalid = int(df["Time"].isna().sum())
    df = df[~kosong & df["Time"].notna()]
    df = df.sort_values(["Case ID", "Time"], kind="stable").reset_index(drop=True)
    info = {"baris_awal": baris_awal, "baris_dipakai": len(df),
            "waktu_invalid": waktu_invalid, "baris_kosong": int(kosong.sum())}
    return df, info


def bangun_trace(df: pd.DataFrame):
    """Langkah 1 & 4. Return (tabel per kontainer, tabel per trace)."""
    g = df.groupby("Case ID", sort=False)
    kasus = pd.DataFrame({
        "Jalur": g["Activity"].agg(PANAH.join),
        "Jumlah Aktivitas": g["Activity"].size(),
        "Mulai": g["Time"].min(),
        "Selesai": g["Time"].max(),
    })
    hitung = kasus["Jalur"].value_counts()
    tabel = pd.DataFrame({"Jalur": hitung.index, "Jumlah Kontainer": hitung.values})
    tabel = tabel.sort_values(["Jumlah Kontainer", "Jalur"], ascending=[False, True],
                              kind="stable").reset_index(drop=True)
    tabel.insert(0, "Trace", [f"Trace {i}" for i in range(1, len(tabel) + 1)])
    tabel["Persen (%)"] = (tabel["Jumlah Kontainer"] / len(kasus) * 100).round(2)
    tabel["Jumlah Aktivitas"] = tabel["Jalur"].str.count(PANAH) + 1
    kasus["Trace"] = kasus["Jalur"].map(dict(zip(tabel["Jalur"], tabel["Trace"])))
    return kasus.reset_index(), tabel


def tabel_aktivitas(df: pd.DataFrame) -> pd.DataFrame:
    """Langkah 3: aktivitas unik beserta frekuensinya."""
    a = df.groupby("Activity").agg(**{
        "Jumlah Kejadian": ("Case ID", "size"),
        "Jumlah Kontainer": ("Case ID", "nunique"),
    }).reset_index()
    a = a.sort_values(["Jumlah Kejadian", "Activity"], ascending=[False, True])
    a.insert(0, "No", range(1, len(a) + 1))
    return a


def trace_per_periode(kasus: pd.DataFrame, bulan: int = 3, basis: str = "Mulai"):
    """Langkah 5: jumlah kontainer per trace untuk tiap blok `bulan` bulan.
    Blok dihitung mulai dari bulan pertama data."""
    t = kasus[basis]
    awal = t.min().to_period("M")
    idx = (t.dt.year - awal.year) * 12 + (t.dt.month - awal.month)
    blok = (idx // bulan).astype(int)
    label = {}
    for b in sorted(blok.unique()):
        s = awal + int(b) * bulan
        label[b] = f"P{b + 1}: {s} s/d {s + bulan - 1}"
    kolom = blok.map(label)
    ct = pd.crosstab(kasus["Trace"], kolom)
    ct = ct[[label[b] for b in sorted(label)]]
    ct = ct.loc[sorted(ct.index, key=lambda x: int(x.split()[1]))]
    ct.columns.name = None
    ct["Total"] = ct.sum(axis=1)
    return ct.reset_index()