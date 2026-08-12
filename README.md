# Allegro → EMPIK

Jedna aplikacja zamiast czterech kroków. Zastępuje cały dotychczasowy łańcuch:

```
stary:  scalarka XLSM  →  upload do repo (input/)  →  GitHub Actions  →  shoperXML  →  XLSX
nowy:   wgraj pliki    →  pobierz XLSX
```

Nikt poza Tobą nie potrzebuje konta ani hasła do GitHuba. Druga osoba dostaje
sam link do aplikacji.

## Co robi pod spodem

1. **Scala** wgrane pliki Allegro — arkusz `Szablon`, nagłówki w wierszu 4,
   kolejność kolumn z pierwszego pliku (identycznie jak dotychczasowa scalarka).
2. **Wygasza wskazane oferty** — jeśli wkleisz numery ofert w pole
   „Oferty do wygaszenia", dostają `Status oferty = Zakończona` i `0` sztuk.
   To odpowiednik makra VBA `OznaczZakonczone`, które szło na pliku przed
   wrzuceniem go do `input/`. Różnica: trafia w kolumny po **nagłówkach**,
   a nie po literach C/H/M, więc dołożenie kolumny przez Allegro nic nie psuje.
3. **Konwertuje** przez `convert_empi.py` → `empi.xml`. To ten sam plik, który
   wcześniej generowały GitHub Actions: dostępność tylko od 4 szt., kategorie
   z dopiskiem „poleasingowe", stan „Używany" → „Poleasingowy", dyski SSD/HDD
   z pojemnością, gwarancja jako liczba, opis w HTML ze stopką Kompre.pl.
4. **Rozbija XML na tabelę** i filtruje — domyślnie kategoria z laptopami,
   status „Wszystkie" (czyli to, co ustawiałeś ręcznie w shoperXML).
5. **Oddaje XLSX** (plus CSV i surowy `empi.xml`, gdyby był potrzebny).

Zgodność sprawdzona na prawdziwych danych: ten sam plik wejściowy przepuszczony
przez aplikację daje `empi.xml` **identyczny co do bajtu** (10 170 895 B,
2248 ofert) z tym, który wyprodukowały GitHub Actions.

Nic nie jest nigdzie zapisywane — pliki żyją tylko przez czas przetwarzania,
w katalogu tymczasowym, który jest kasowany po zakończeniu.

## Pliki

| Plik | Skąd |
|---|---|
| `app.py` | interfejs (nowy) |
| `pipeline.py` | scalanie + spięcie kroków + XML→tabela (nowy) |
| `convert.py` | **kopia 1:1** z `nowe_repo10.2025_allegrocsv_na_XML/scripts/` |
| `convert_empi.py` | **kopia 1:1** z `nowe_repo10.2025_allegrocsv_na_XML/scripts/` |

`convert.py` i `convert_empi.py` są celowo **niezmienione**, żeby wynik był
identyczny z tym, co produkuje workflow. Jeśli poprawisz reguły w oryginalnym
repo, po prostu skopiuj oba pliki tutaj ponownie — reszta nie wymaga zmian.

## Uruchomienie u siebie (lokalnie)

```bash
pip install -r requirements.txt
```

```bash
streamlit run app.py
```

## Wystawienie dla drugiej osoby (raz, ~10 minut)

1. Załóż **nowe, publiczne** repo na GitHubie, np. `allegro-empik`.
2. Wrzuć do jego **głównego katalogu** zawartość tego folderu
   (`app.py`, `pipeline.py`, `convert.py`, `convert_empi.py`, `requirements.txt`).
3. Wejdź na [share.streamlit.io](https://share.streamlit.io) → **New app**,
   wskaż to repo, gałąź `main`, plik główny `app.py` → **Deploy**.
4. W ustawieniach aplikacji (⚙️ → *Sharing*) ustaw dostęp na **publiczny**
   („anyone with the link"), inaczej druga osoba zobaczy ekran logowania —
   tak jak dziś w scalarce i shoperXML.
5. Wyślij jej link. Koniec.

W repo nie ma żadnych danych ani sekretów — sam kod przetwarzający. Osoba
korzystająca z aplikacji widzi wyłącznie to, co sama wgrała.

## Jeśli chcesz dodatkowo zabezpieczyć aplikację hasłem

Streamlit Community Cloud nie ma prostego hasła do publicznej aplikacji,
ale można je dorobić w kodzie: dodaj w panelu aplikacji sekret

```toml
haslo = "wpisz-cos-tutaj"
```

i na początku `app.py`:

```python
if st.text_input("Hasło", type="password") != st.secrets["haslo"]:
    st.stop()
```

To hasło do aplikacji, nie do Twojego konta — możesz je swobodnie przekazać
i w każdej chwili zmienić.

## Stary proces

Nie musisz go kasować. Repo `nowe_repo10.2025_allegrocsv_na_XML`, workflow
i shoperXML działają dalej bez zmian — ta aplikacja niczego tam nie dotyka.
