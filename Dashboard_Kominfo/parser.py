"""
parser.py
---------
Modul untuk membaca & membersihkan data hasil rekap CCTV lalu lintas.

Mendukung dua skema:
1. Skema "blok per lokasi" (format khas rekap CCTV):
   - Setiap lokasi memiliki baris header sendiri (kolom ke-2 = "Waktu"),
     diikuti baris-baris data harian, lalu baris "Total", lalu baris kosong.
   - Urutan & jumlah kolom kategori (Orang/Motor/Mobil/Truk/Bus) bisa
     berbeda-beda tiap lokasi -> dibaca dinamis dari baris header masing2 blok.
2. Skema generik/tabular biasa (satu header, satu tabel lurus) sebagai fallback,
   untuk berjaga-jaga jika user upload file dengan format yang lebih sederhana.

Hasil akhir selalu berupa DataFrame "tidy" dengan kolom:
    Lokasi, Tanggal, Orang, Motor, Mobil, Truk, Bus, Total
"""

import io
import pandas as pd
import numpy as np

CATEGORIES = ["Orang", "Motor", "Mobil", "Truk", "Bus"]
ALL_NUMERIC_COLS = CATEGORIES + ["Total"]

BULAN_ID = [
    "",
    "Januari",
    "Februari",
    "Maret",
    "April",
    "Mei",
    "Juni",
    "Juli",
    "Agustus",
    "September",
    "Oktober",
    "November",
    "Desember",
]

HARI_ID = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]


def format_tanggal_id(dt, with_year=True, with_time=False) -> str:
    """Format objek datetime/Timestamp menjadi tanggal berbahasa Indonesia, mis. '01 Januari 2026'."""
    if pd.isna(dt):
        return "-"
    s = f"{dt.day:02d} {BULAN_ID[dt.month]}"
    if with_year:
        s += f" {dt.year}"
    if with_time:
        s += f", {dt.strftime('%H:%M')}"
    return s


class DataFormatError(Exception):
    """Dilempar ketika file tidak bisa dikenali formatnya sama sekali."""

    pass


def _read_raw(uploaded_file):
    """Membaca file upload (csv/xlsx/xls) menjadi DataFrame mentah tanpa header."""
    name = getattr(uploaded_file, "name", "") or ""
    ext = name.split(".")[-1].lower()

    # Pastikan pointer file di awal (Streamlit UploadedFile bisa dibaca ulang)
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)

    if ext in ("xlsx", "xls", "xlsm"):
        raw = pd.read_excel(uploaded_file, header=None, dtype=object)
    elif ext == "csv":
        raw = pd.read_csv(
            uploaded_file, header=None, dtype=object, sep=None, engine="python"
        )
    else:
        # coba sebagai excel dulu, lalu csv
        try:
            if hasattr(uploaded_file, "seek"):
                uploaded_file.seek(0)
            raw = pd.read_excel(uploaded_file, header=None, dtype=object)
        except Exception:
            if hasattr(uploaded_file, "seek"):
                uploaded_file.seek(0)
            raw = pd.read_csv(
                uploaded_file, header=None, dtype=object, sep=None, engine="python"
            )
    return raw


def _parse_block_format(raw: pd.DataFrame) -> pd.DataFrame:
    """Parser untuk format blok per lokasi."""
    records = []
    col_map = {}
    expected_categories = []

    n_rows, n_cols = raw.shape

    for i in range(n_rows):
        row = raw.iloc[i]

        if row.isna().all():
            continue

        val1 = row[1] if n_cols > 1 else None

        # Header setiap blok lokasi
        if isinstance(val1, str) and val1.strip().lower() == "waktu":
            col_map = {}

            for idx in range(2, n_cols):
                v = row[idx]
                if isinstance(v, str) and v.strip():
                    col_map[idx] = v.strip().capitalize()

            # Beberapa blok Excel tidak menyimpan sel kosong untuk kategori
            # "Orang". Akibatnya, header blok dapat bergeser satu kolom ke kiri,
            # sementara nilai data tetap berada pada posisi kolom aslinya.
            # Ciri kasus ini: Motor berada di kolom pertama kategori (indeks 2),
            # Total berada di indeks 6, dan kolom terakhir (indeks 7) kosong.
            if (
                col_map.get(2) == "Motor"
                and col_map.get(6) == "Total"
                and n_cols > 7
                and pd.isna(row[7])
                and "Orang" not in col_map.values()
            ):
                col_map = {idx + 1: label for idx, label in col_map.items()}

            expected_categories = [
                label for label in col_map.values() if label in CATEGORIES
            ]
            continue

        # Lewati baris ringkasan Total
        if isinstance(val1, str) and val1.strip().lower() == "total":
            continue

        tanggal = pd.to_datetime(val1, errors="coerce")
        if pd.isna(tanggal):
            continue

        lokasi = row[0]
        if pd.isna(lokasi) or str(lokasi).strip().lower() == "lokasi":
            continue

        rec = {
            "Lokasi": str(lokasi).strip(),
            "Tanggal": tanggal,
        }

        for cat in CATEGORIES:
            rec[cat] = np.nan

        total_raw = np.nan

        for idx, label in col_map.items():
            if idx >= n_cols:
                continue

            value = pd.to_numeric(row[idx], errors="coerce")

            if label in CATEGORIES:
                rec[label] = value
            elif label.lower() == "total":
                total_raw = value

        # Hanya kategori yang tercantum di header blok
        # yang digunakan untuk pemeriksaan kelengkapan.
        expected_values = [rec[c] for c in expected_categories]

        rec["_expected_count"] = len(expected_categories)
        rec["_valid_count"] = sum(pd.notna(v) for v in expected_values)
        rec["_missing_count"] = rec["_expected_count"] - rec["_valid_count"]

        computed_total = (
            pd.Series(expected_values, dtype="float64").sum(min_count=1)
            if expected_values
            else np.nan
        )

        # Pertahankan Total sumber jika tersedia.
        rec["Total"] = total_raw if pd.notna(total_raw) else computed_total

        # Periksa kecocokan hanya jika seluruh kategori yang
        # diharapkan tersedia dan Total dapat dibandingkan.
        rec["_total_match"] = bool(
            rec["_valid_count"] == rec["_expected_count"]
            and pd.notna(rec["Total"])
            and pd.notna(computed_total)
            and np.isclose(
                rec["Total"],
                computed_total,
                rtol=1e-5,
                atol=1e-8,
            )
        )

        records.append(rec)

    if not records:
        return None

    df = pd.DataFrame(records)

    # Pertahankan nilai kosong untuk pemeriksaan kualitas.
    for col in ALL_NUMERIC_COLS:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


def _parse_generic_format(uploaded_file) -> pd.DataFrame:
    """Fallback: file tabular biasa dengan satu baris header."""
    name = getattr(uploaded_file, "name", "") or ""
    ext = name.split(".")[-1].lower()

    # Pastikan file dibaca dari posisi awal
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)

    # Baca file Excel atau CSV
    if ext in ("xlsx", "xls", "xlsm"):
        df = pd.read_excel(uploaded_file)
    else:
        df = pd.read_csv(uploaded_file, sep=None, engine="python")

    # Rapikan nama kolom
    df.columns = [str(c).strip() for c in df.columns]
    colmap_lower = {c.lower(): c for c in df.columns}

    def find_col(*names):
        for name in names:
            if name.lower() in colmap_lower:
                return colmap_lower[name.lower()]
        return None

    # Cari kolom lokasi dan tanggal
    loc_col = find_col("lokasi", "location", "nama lokasi", "titik", "site")
    date_col = find_col("tanggal", "waktu", "date", "day", "hari")

    if loc_col is None or date_col is None:
        raise DataFormatError(
            "Tidak ditemukan kolom Lokasi dan/atau Tanggal/Waktu pada file. "
            "Pastikan file memiliki kolom lokasi dan tanggal yang jelas."
        )

    # Samakan nama kolom
    rename = {
        loc_col: "Lokasi",
        date_col: "Tanggal",
    }

    for cat in CATEGORIES:
        col = find_col(cat)
        if col is not None:
            rename[col] = cat

    total_col = find_col("total")
    if total_col is not None:
        rename[total_col] = "Total"

    df = df.rename(columns=rename)

    # Konversi tanggal dan hapus baris tanpa identitas lokasi/tanggal
    df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors="coerce")
    df = df.dropna(subset=["Tanggal", "Lokasi"]).copy()
    df["Lokasi"] = df["Lokasi"].astype(str).str.strip()

    # Pastikan kolom kategori tersedia.
    # Nilai kosong tetap NaN, JANGAN diubah menjadi 0 di tahap ini.
    for cat in CATEGORIES:
        if cat not in df.columns:
            df[cat] = np.nan
        else:
            df[cat] = pd.to_numeric(df[cat], errors="coerce")

    # Jumlah kategori yang seharusnya tersedia pada format tabular ini
    df["_expected_count"] = len(CATEGORIES)

    # Tandai jumlah kategori kosong dan kategori yang memiliki nilai
    df["_missing_count"] = df[CATEGORIES].isna().sum(axis=1)
    df["_valid_count"] = df[CATEGORIES].notna().sum(axis=1)

    # Hitung total kategori yang tersedia.
    # min_count=1 menjaga hasil tetap NaN jika seluruh kategori kosong.
    computed_total = df[CATEGORIES].sum(axis=1, min_count=1)

    # Gunakan Total dari sumber jika tersedia.
    # Jika Total sumber kosong, gunakan hasil penjumlahan kategori.
    if "Total" not in df.columns:
        df["Total"] = computed_total
    else:
        df["Total"] = pd.to_numeric(df["Total"], errors="coerce")
        df["Total"] = df["Total"].where(df["Total"].notna(), computed_total)

    # Tandai apakah Total cocok dengan penjumlahan kategori.
    # Baris dengan kategori yang belum lengkap tidak dinyatakan cocok.
    all_categories_valid = df["_valid_count"] == df["_expected_count"]

    df["_total_match"] = (
        all_categories_valid
        & df["Total"].notna()
        & computed_total.notna()
        & np.isclose(
            df["Total"].fillna(0), computed_total.fillna(0), rtol=1e-5, atol=1e-8
        )
    )

    # Kembalikan kolom utama beserta penanda kualitas data
    return df[
        ["Lokasi", "Tanggal"]
        + ALL_NUMERIC_COLS
        + [
            "_expected_count",
            "_missing_count",
            "_valid_count",
            "_total_match",
        ]
    ]


def load_and_clean(uploaded_file) -> pd.DataFrame:
    """Membaca dan membersihkan data CCTV."""
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)

    raw = None

    try:
        name = getattr(uploaded_file, "name", "") or ""
        ext = name.split(".")[-1].lower()

        if ext in ("xlsx", "xls", "xlsm"):
            raw = pd.read_excel(uploaded_file, header=None)
        else:
            raw = pd.read_csv(uploaded_file, sep=None, engine="python", header=None)

        df = _parse_block_format(raw)

        if df is not None and not df.empty:
            return df.sort_values(["Lokasi", "Tanggal"]).reset_index(drop=True)

    except Exception:
        pass

    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)

    df = _parse_generic_format(uploaded_file)

    return df.sort_values(["Lokasi", "Tanggal"]).reset_index(drop=True)


def data_quality_report(df):
    """Ringkasan kualitas data CCTV per lokasi."""
    df = df.copy()

    # Pastikan kolom kategori tersedia.
    for cat in CATEGORIES:
        if cat not in df.columns:
            df[cat] = np.nan
        else:
            df[cat] = pd.to_numeric(df[cat], errors="coerce")

    # Pastikan Total tersedia dan tetap mempertahankan Total sumber.
    if "Total" not in df.columns:
        df["Total"] = df[CATEGORIES].sum(axis=1, min_count=1)
    else:
        df["Total"] = pd.to_numeric(df["Total"], errors="coerce")

    # Penanda kualitas: pertahankan dari parser jika sudah tersedia.
    if "_expected_count" not in df.columns:
        df["_expected_count"] = len(CATEGORIES)

    if "_valid_count" not in df.columns:
        df["_valid_count"] = df[CATEGORIES].notna().sum(axis=1)

    if "_missing_count" not in df.columns:
        df["_missing_count"] = (df["_expected_count"] - df["_valid_count"]).clip(
            lower=0
        )

    # Periksa kecocokan Total dengan penjumlahan kategori.
    computed_total = df[CATEGORIES].sum(axis=1, min_count=1)
    all_valid = df["_valid_count"] == df["_expected_count"]

    calculated_match = (
        all_valid
        & df["Total"].notna()
        & computed_total.notna()
        & np.isclose(
            df["Total"].fillna(0),
            computed_total.fillna(0),
            rtol=1e-5,
            atol=1e-8,
        )
    )

    if "_total_match" not in df.columns:
        df["_total_match"] = calculated_match
    else:
        df["_total_match"] = (
            df["_total_match"].fillna(False).astype(bool) & calculated_match
        )

    df["_total_mismatch"] = all_valid & ~df["_total_match"]

    # Satu baris per lokasi dan tanggal agar jumlah hari tidak terhitung
    # ganda jika terdapat duplikasi tanggal.
    harian = (
        df.sort_values("Tanggal")
        .groupby(["Lokasi", "Tanggal"], as_index=False)
        .agg(
            Total=("Total", lambda s: s.sum(min_count=1)),
            missing_count=("_missing_count", "max"),
            valid_count=("_valid_count", "max"),
            expected_count=("_expected_count", "max"),
            total_mismatch=("_total_mismatch", "max"),
        )
    )

    summary = (
        harian.groupby("Lokasi")
        .agg(
            jumlah_hari_tercatat=("Tanggal", "nunique"),
            total_keseluruhan=("Total", lambda s: s.sum(min_count=1)),
            rata_rata_harian=("Total", "mean"),
            hari_nol=("Total", lambda s: (s == 0).sum()),
            hari_data_kosong=("missing_count", lambda s: (s > 0).sum()),
            hari_tanpa_data_valid=("valid_count", lambda s: (s == 0).sum()),
            hari_total_tidak_cocok=("total_mismatch", "sum"),
        )
        .reset_index()
    )

    def tentukan_status(row):
        total_hari = row["jumlah_hari_tercatat"]
        hari_nol = row["hari_nol"]
        hari_kosong = row["hari_data_kosong"]
        hari_tanpa_data = row["hari_tanpa_data_valid"]
        hari_tidak_cocok = row["hari_total_tidak_cocok"]

        if total_hari == 0:
            return "Tidak ada data / kamera tidak aktif"

        if hari_tanpa_data == total_hari:
            return "Tidak ada data / kamera tidak aktif"

        # Jika seluruh hari Total-nya nol, tandai untuk pemeriksaan.
        if hari_nol == total_hari:
            return "Tidak ada data / kamera tidak aktif"

        # Ketidakcocokan Total harus diperiksa.
        if hari_tidak_cocok > 0:
            return "Total tidak cocok"

        # Data kosong harus ditandai, terlepas dari jumlah hari nol.
        if hari_kosong > 0:
            return "Sebagian data kosong"

        # Banyak hari nol tidak otomatis berarti data tidak ada.
        # Jika lebih dari 30% hari bernilai nol, perlu verifikasi.
        if hari_nol / total_hari > 0.30:
            return "Perlu verifikasi nilai nol"

        return "Data lengkap"

    summary["status"] = summary.apply(tentukan_status, axis=1)
    summary["rata_rata_harian"] = summary["rata_rata_harian"].round(0).astype("Int64")

    return summary.sort_values(
        "total_keseluruhan",
        ascending=False,
        na_position="last",
    ).reset_index(drop=True)


def aggregate_by_period(df: pd.DataFrame, period: str) -> pd.DataFrame:
    """
    Agregasi data tidy berdasarkan periode: 'Harian', 'Mingguan', 'Bulanan', 'Keseluruhan'.
    Mengembalikan DataFrame dengan kolom Periode, Lokasi, Orang..Total.
    """
    df = df.copy()
    if period == "Harian":
        df["Periode"] = df["Tanggal"].dt.strftime("%Y-%m-%d")
    elif period == "Mingguan":
        iso = df["Tanggal"].dt.isocalendar()
        df["Periode"] = (
            iso["year"].astype(str) + "-W" + iso["week"].astype(str).str.zfill(2)
        )
    elif period == "Bulanan":
        df["Periode"] = df["Tanggal"].dt.strftime("%Y-%m")
    else:  # Keseluruhan
        lo = df["Tanggal"].min().strftime("%d %b %Y")
        hi = df["Tanggal"].max().strftime("%d %b %Y")
        df["Periode"] = f"{lo} – {hi}"

    agg = df.groupby(["Periode", "Lokasi"], as_index=False)[ALL_NUMERIC_COLS].sum()
    return agg
