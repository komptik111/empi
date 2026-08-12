# -*- coding: utf-8 -*-
"""
Cały łańcuch Allegro -> EMPIK w jednym miejscu, bez GitHuba i bez Actions.

Kroki (dokładnie te same, co w dotychczasowym procesie):
  1. scalenie plików XLSM/XLSX z arkusza "Szablon" (nagłówki w wierszu 4)
  2. konwersja przez ORYGINALNE scripts/convert_empi.py -> empi.xml
  3. rozbicie XML z powrotem na tabelę (tak jak robi to shoperXML)

convert.py i convert_empi.py są skopiowane 1:1 z repo
marekkomp/nowe_repo10.2025_allegrocsv_na_XML — nie modyfikujemy ich,
żeby wynik był identyczny z tym, co produkuje workflow.
"""

import os
import re
import tempfile
import threading
import xml.etree.ElementTree as ET

import pandas as pd

SHEET_NAME = "Szablon"
HEADER_ROW_IDX = 3  # nagłówki w wierszu 4 (0-based -> 3)

# Odpowiednik makra VBA OznaczZakonczone: kolumny C / H / M w szablonie Allegro.
COL_ID = "ID oferty"
COL_STATUS = "Status oferty"
COL_QTY = "Liczba sztuk"
ENDED_STATUS = "Zakończona"

# convert_empi trzyma ścieżkę roboczą w module-level OUTPUT_DIR i zapisuje tam
# _temp_base.xml. Podmieniamy ją na katalog tymczasowy, więc konwersje nie mogą
# wejść sobie w drogę — lock pilnuje, żeby dwie naraz nie nadpisały tej zmiennej.
_convert_lock = threading.Lock()


class PipelineError(Exception):
    """Błąd, który ma sens pokazać użytkownikowi wprost."""


# --------------------------------------------------------------------------
# 1. Scalanie plików
# --------------------------------------------------------------------------
def read_allegro_file(file_obj, filename: str) -> pd.DataFrame:
    """Wczytuje jeden plik Allegro: arkusz 'Szablon', nagłówki z wiersza 4."""
    try:
        df = pd.read_excel(
            file_obj, sheet_name=SHEET_NAME, header=HEADER_ROW_IDX, dtype=str
        )
    except ValueError:
        # brak arkusza "Szablon" — bierzemy pierwszy, tak jak robi to convert.py
        file_obj.seek(0)
        df = pd.read_excel(file_obj, sheet_name=0, header=HEADER_ROW_IDX, dtype=str)
    except Exception as e:
        raise PipelineError(f"Nie udało się wczytać pliku {filename}: {e}") from e

    return df.fillna("")


def merge_frames(frames: list[pd.DataFrame]) -> pd.DataFrame:
    """
    Scala po NAZWACH nagłówków.
    Kolejność kolumn = kolejność z pierwszego pliku, nowe kolumny dopisywane na końcu.
    """
    if not frames:
        raise PipelineError("Brak danych do scalenia.")

    column_order: list[str] = []
    for df in frames:
        for c in df.columns:
            if c not in column_order:
                column_order.append(c)

    prepared = []
    for df in frames:
        for missing in [c for c in column_order if c not in df.columns]:
            df[missing] = ""
        prepared.append(df[column_order])

    return pd.concat(prepared, axis=0, ignore_index=True)


def parse_ids(text: str) -> list[str]:
    """
    Wyciąga numery ofert z wklejonego tekstu — obojętne, czy rozdzielone
    przecinkami, spacjami, czy każdy w nowej linii. Duplikaty pomijane.
    """
    if not text:
        return []
    out, seen = [], set()
    for part in re.split(r"\D+", text):
        if part and part not in seen:
            seen.add(part)
            out.append(part)
    return out


def mark_ended(
    merged: pd.DataFrame, ids: list[str]
) -> tuple[pd.DataFrame, list[str], list[str]]:
    """
    To samo, co makro VBA OznaczZakonczone: dla podanych ID ustawia
    status "Zakończona" i 0 sztuk. Reszta łańcucha sama wygasi te oferty.

    Zwraca (tabela, znalezione ID, ID nieobecne w plikach).
    """
    if not ids:
        return merged, [], []

    brakujace = [c for c in (COL_ID, COL_STATUS, COL_QTY) if c not in merged.columns]
    if brakujace:
        raise PipelineError(
            "W plikach brakuje kolumn potrzebnych do wygaszania ofert: "
            + ", ".join(brakujace)
        )

    key = merged[COL_ID].astype(str).str.strip()
    mask = key.isin(set(ids))

    merged = merged.copy()
    merged.loc[mask, COL_STATUS] = ENDED_STATUS
    merged.loc[mask, COL_QTY] = "0"

    znalezione = set(key[mask])
    return (
        merged,
        [i for i in ids if i in znalezione],
        [i for i in ids if i not in znalezione],
    )


def write_template_xlsx(merged: pd.DataFrame, path: str) -> None:
    """Zapisuje scalony arkusz w formacie, którego oczekuje convert.py."""
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        merged.to_excel(
            writer, index=False, sheet_name=SHEET_NAME, startrow=HEADER_ROW_IDX
        )


# --------------------------------------------------------------------------
# 2. Konwersja do empi.xml
# --------------------------------------------------------------------------
def build_empi_xml(merged: pd.DataFrame) -> bytes:
    """Odpala oryginalny convert_empi.convert_file_empi() na scalonym pliku."""
    import convert_empi

    # ignore_cleanup_errors: convert.py nie zamyka workbooka w ścieżce sukcesu,
    # więc na Windows uchwyt do pliku blokuje kasowanie katalogu. Na Linuksie
    # (Streamlit Cloud) problem nie występuje, ale nie chcemy wywalki lokalnie.
    with tempfile.TemporaryDirectory(prefix="empi_", ignore_cleanup_errors=True) as workdir:
        src = os.path.join(workdir, "allegro_merged.xlsx")
        dst = os.path.join(workdir, "empi.xml")
        write_template_xlsx(merged, src)

        with _convert_lock:
            original_output_dir = convert_empi.OUTPUT_DIR
            convert_empi.OUTPUT_DIR = workdir
            try:
                convert_empi.convert_file_empi(src, dst)
            finally:
                convert_empi.OUTPUT_DIR = original_output_dir

        if not os.path.exists(dst):
            raise PipelineError("Konwerter nie wygenerował pliku XML.")
        with open(dst, "rb") as f:
            return f.read()


# --------------------------------------------------------------------------
# 3. XML -> tabela (logika przeniesiona z shoperXML/app.py)
# --------------------------------------------------------------------------
BASE_COLUMNS = [
    "Kategoria",
    "Podkategoria",
    "Producent",
    "Nazwa",
    "Cena",
    "Dostępność",
    "Liczba sztuk",
    "ID",
    "URL",
    "Opis HTML",
]


def xml_to_dataframe(xml_bytes: bytes) -> pd.DataFrame:
    root = ET.fromstring(xml_bytes)

    rows = []
    max_imgs = 0

    for o in root.findall(".//o"):
        desc_html = ""
        desc_el = o.find("desc")
        if desc_el is not None:
            desc_html = "".join(
                ET.tostring(child, encoding="unicode", method="xml")
                for child in list(desc_el)
            ).strip() or (desc_el.text or "").strip()

        images = []
        imgs_el = o.find("imgs")
        if imgs_el is not None:
            main_el = imgs_el.find("main")
            if main_el is not None and (main_el.get("url") or "").strip():
                images.append(main_el.get("url").strip())
            for i_el in imgs_el.findall("i"):
                u = (i_el.get("url") or "").strip()
                if u:
                    images.append(u)
        max_imgs = max(max_imgs, len(images))

        producent = ""
        extra = {}
        attrs_el = o.find("attrs")
        if attrs_el is not None:
            for a in attrs_el.findall("a"):
                k = (a.get("name") or "").strip()
                if not k:
                    continue
                v = (a.text or "").strip()
                extra[k] = v
                if k.lower() == "producent":
                    producent = v

        avail = (o.get("avail") or "").strip()
        row = {
            "Kategoria": (o.findtext("cat") or "").strip(),
            "Podkategoria": (o.findtext("subcat") or "").strip(),
            "Producent": producent,
            "Nazwa": (o.findtext("name") or "").strip(),
            "Cena": (o.get("price") or "").strip().replace(",", "."),
            "Dostępność": 1 if avail in {"1", "true", "True", "tak", "TAK"} else 99,
            "Liczba sztuk": (o.get("stock") or "").strip(),
            "ID": (o.get("id") or "").strip(),
            "URL": (o.get("url") or "").strip(),
            "Opis HTML": desc_html,
        }

        for i, img in enumerate(images):
            row[f"Zdjęcie {i + 1}"] = img
        for k, v in extra.items():
            row.setdefault(k, v)

        rows.append(row)

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    for i in range(1, max_imgs + 1):
        col = f"Zdjęcie {i}"
        if col not in df.columns:
            df[col] = ""

    for c in ("Cena", "Dostępność", "Liczba sztuk"):
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    return df


def drop_empty_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Ukrywa kolumny bez ani jednej wartości — tak jak 'widok niepuste' w shoperXML."""
    keep = [
        c
        for c in df.columns
        if (df[c].notna() & ~df[c].astype(str).str.strip().eq("")).any()
    ]
    return df[keep]


def laptop_categories(df: pd.DataFrame) -> list[str]:
    """
    Kategorie z samymi laptopami — po convert_empi to 'Laptopy poleasingowe'.

    Dopasowanie po POCZĄTKU nazwy, nie po fragmencie. Szukanie 'laptop'
    w środku łapałoby też 'Części do laptopów' i 'Akcesoria (Laptop, PC)',
    a te do EMPIK nie idą.
    """
    if "Kategoria" not in df.columns:
        return []
    cats = df["Kategoria"].dropna().astype(str).str.strip().unique().tolist()
    return sorted([c for c in cats if c.lower().startswith("laptop")])


# --------------------------------------------------------------------------
# Całość w jednym wywołaniu
# --------------------------------------------------------------------------
def run_pipeline(uploaded_files, ended_ids=()) -> tuple[bytes, pd.DataFrame, dict]:
    """
    uploaded_files: lista obiektów ze Streamlita (mają .name i są file-like).
    ended_ids: numery ofert do wygaszenia (zastępuje makro VBA).

    Zwraca (empi.xml jako bytes, tabela ofert, informacje o wygaszaniu).
    """
    frames = []
    for f in uploaded_files:
        f.seek(0)
        frames.append(read_allegro_file(f, f.name))

    merged = merge_frames(frames)
    if merged.empty:
        raise PipelineError("Po scaleniu nie ma żadnych wierszy z ofertami.")

    merged, wygaszone, nieznalezione = mark_ended(merged, list(ended_ids))
    info = {"wygaszone": wygaszone, "nieznalezione": nieznalezione}

    xml_bytes = build_empi_xml(merged)
    df = xml_to_dataframe(xml_bytes)
    if df.empty:
        raise PipelineError(
            "Konwerter nie znalazł ofert. Sprawdź, czy pliki mają arkusz "
            "'Szablon' z nagłówkami w wierszu 4 (m.in. 'ID oferty', 'Tytuł oferty')."
        )
    return xml_bytes, df, info
