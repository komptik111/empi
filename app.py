# -*- coding: utf-8 -*-
"""
Allegro -> EMPIK w jednym kroku.

Zastępuje łańcuch: scalarka -> upload do repo -> GitHub Actions -> shoperXML.
Nikt nie potrzebuje konta ani hasła do GitHuba — wgrywa pliki i pobiera wynik.
"""

from io import BytesIO

import pandas as pd
import streamlit as st

from pipeline import (
    PipelineError,
    drop_empty_columns,
    laptop_categories,
    parse_ids,
    run_pipeline,
)

st.set_page_config(page_title="Allegro → EMPIK", layout="wide", page_icon="💻")

st.title("💻 Allegro → EMPIK")
st.caption(
    "Wgraj pliki Excel pobrane z Allegro. Aplikacja sama je scali, przekonwertuje "
    "i przygotuje plik z laptopami pod EMPIK."
)


@st.cache_data(show_spinner=False, max_entries=3)
def _process(payload: list[tuple[str, bytes]], ended_ids: tuple[str, ...]):
    """Cache po zawartości plików — ponowne kliknięcie nie liczy wszystkiego od nowa."""
    files = []
    for name, data in payload:
        buf = BytesIO(data)
        buf.name = name
        files.append(buf)
    return run_pipeline(files, ended_ids)


@st.cache_data(show_spinner=False)
def _to_xlsx(df: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="dane")
    return out.getvalue()


uploaded = st.file_uploader(
    "Pliki z Allegro (XLSM / XLSX / XLS) — możesz wgrać kilka naraz",
    type=["xlsm", "xlsx", "xls"],
    accept_multiple_files=True,
)

if not uploaded:
    with st.expander("Jak tego użyć?", expanded=True):
        st.markdown(
            """
1. Pobierz z Allegro pliki Excel z ofertami (jeden lub kilka).
2. Dodaj jako pierwszy Główne konto!!! potem reszta
3. Jeśli któreś oferty mają być wygaszone — wklej ich numery w pole poniżej.
4. Poczekaj kilkanaście sekund.
5. Kliknij **Pobierz XLSX** — to jest plik pod EMPIK.

Nic nie jest nigdzie zapisywane. Pliki żyją tylko przez czas przetwarzania.
            """
        )
    st.stop()

st.write(f"Wgrane pliki: **{len(uploaded)}** — {', '.join(f.name for f in uploaded)}")

with st.expander("Oferty do wygaszenia (opcjonalne)", expanded=False):
    st.caption(
        "Wklej numery ofert, które mają trafić do pliku jako zakończone, 0 sztuk — "
        "niezależnie od tego, co mówi Allegro. Zastępuje makro VBA. "
        "Przecinki, spacje albo każdy numer w nowej linii — bez różnicy."
    )
    ended_raw = st.text_area(
        "Numery ofert",
        value="",
        height=110,
        placeholder="17897523407, 16998056567, 16820940558",
        label_visibility="collapsed",
    )

ended_ids = tuple(parse_ids(ended_raw))
if ended_ids:
    st.caption(f"Do wygaszenia: {len(ended_ids)} numerów.")

payload = [(f.name, f.getvalue()) for f in uploaded]

try:
    with st.spinner("Scalam, konwertuję i przygotowuję oferty…"):
        xml_bytes, df, info = _process(payload, ended_ids)
except PipelineError as e:
    st.error(str(e))
    st.stop()
except Exception as e:  # noqa: BLE001 — użytkownik ma zobaczyć powód, nie stacktrace
    st.error(f"Coś poszło nie tak przy przetwarzaniu: {e}")
    st.stop()

st.success(f"Gotowe. Ofert w pliku: {len(df):,}")

if ended_ids:
    st.info(f"Wygaszono ofert: **{len(info['wygaszone'])}** z {len(ended_ids)} podanych.")
    if info["nieznalezione"]:
        with st.expander(
            f"⚠️ {len(info['nieznalezione'])} numerów nie ma w wgranych plikach"
        ):
            st.caption(
                "Te oferty nie występują w eksporcie z Allegro — najpewniej zostały "
                "już wcześniej usunięte. Możesz je wykreślić ze swojej listy."
            )
            st.code("\n".join(info["nieznalezione"]))

# ---------------------------------------------------------------- filtry
st.sidebar.header("Filtry")

default_cats = laptop_categories(df)
all_cats = sorted(df["Kategoria"].dropna().astype(str).str.strip().unique().tolist())

selected_cats = st.sidebar.multiselect(
    "Kategoria",
    options=all_cats,
    default=default_cats,
    help="Domyślnie zaznaczone są laptopy — tak jak w dotychczasowym procesie.",
)

status_choice = st.sidebar.radio(
    "Status produktu",
    options=["Wszystkie", "Aktywne", "Nieaktywne"],
    index=0,
)

producers = sorted(df["Producent"].dropna().astype(str).str.strip().unique().tolist())
selected_prods = st.sidebar.multiselect("Producent (opcjonalnie)", options=producers)

name_query = st.sidebar.text_input("Szukaj w nazwie", value="")

# ---------------------------------------------------------------- filtrowanie
mask = pd.Series(True, index=df.index)

if selected_cats:
    mask &= df["Kategoria"].astype(str).str.strip().isin(selected_cats)

if status_choice != "Wszystkie":
    target = 1 if status_choice == "Aktywne" else 99
    mask &= pd.to_numeric(df["Dostępność"], errors="coerce") == target

if selected_prods:
    mask &= df["Producent"].astype(str).str.strip().isin(selected_prods)

if name_query.strip():
    mask &= df["Nazwa"].astype(str).str.contains(name_query.strip(), case=False, na=False)

filtered = df.loc[mask]

if filtered.empty:
    st.warning("Brak ofert po zastosowaniu filtrów. Poluzuj filtry po lewej stronie.")
    st.stop()

view_df = drop_empty_columns(filtered)

st.subheader("Wynik")
st.write(
    f"Oferty: **{len(view_df):,}** z {len(df):,} • "
    f"Kolumny (niepuste): **{len(view_df.columns):,}**"
)
st.dataframe(view_df, use_container_width=True, height=520)

# ---------------------------------------------------------------- pobieranie
st.divider()
st.subheader("Pobierz")

c1, c2, c3 = st.columns(3)
with c1:
    st.download_button(
        "⬇️ Pobierz XLSX (pod EMPIK)",
        _to_xlsx(view_df),
        "oferty_empik.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
with c2:
    st.download_button(
        "⬇️ CSV",
        view_df.to_csv(index=False).encode("utf-8-sig"),
        "oferty_empik.csv",
        "text/csv",
        use_container_width=True,
    )
with c3:
    st.download_button(
        "⬇️ empi.xml",
        xml_bytes,
        "empi.xml",
        "application/xml",
        help="Ten sam plik, który wcześniej generowały GitHub Actions.",
        use_container_width=True,
    )

st.caption("Kolumny bez żadnej wartości są ukrywane — tak samo jak w shoperXML.")
