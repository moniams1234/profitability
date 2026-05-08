# Postkalkulacja Profitability

Aplikacja Streamlit do przygotowania postkalkulacji rentowności zleceń i wygenerowania pliku XLSX.

## Uruchomienie lokalne

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Wdrożenie na Streamlit Cloud

1. Wrzuć pliki projektu do repozytorium GitHub.
2. W Streamlit Cloud wskaż repozytorium i plik `app.py`.
3. Aplikacja zainstaluje biblioteki z `requirements.txt`.

## Pliki wejściowe

Wymagany:
- Baza / post_list

Opcjonalne:
- Czasy dla aplikacji
- Zlecenia + faktury
- Faktury linie faktury
- Kliki / inks
- Koszty klików
- Stawki dla aplikacji
- Farby podsumowanie

Aplikacja działa także przy braku plików opcjonalnych i pokazuje ostrzeżenia.

## Najważniejsze obliczenia

- `Zlecenie produkcyjne` — część kolumny `Numer` przed znakiem `-`.
- `Lewy 10` — pierwsze 10 znaków z kolumny `Zamówienie`.
- `Batch` — przedział ilości na podstawie kolumny `Zamawiana ilość`.
- `Digital/Offset` — klasyfikacja na podstawie maszyn z arkusza `czasy`.
- `Sales Value` — najpierw z arkusza `zlec + faktury`, a jeśli brak, fallback z `faktury linie`.
- Fallback dla Kohlpharma i podobnych przypadków: z zamówienia typu `KOH-34440-K659995 | 260220_Digital [K659995]` pobierany jest fragment `260220_Digital`, a następnie szukany jako część tekstu w `Nazwa linii faktury`.
- `Total DL` — koszty maszyn + Prepress costs.
- `Total Materials` — Papier, Klej, Lakiery, Opakowania zbiorcze, Other Materials, Offset inks, Płyta offsetowa, Kliki final.
- `TPM` = Sales Value - Total DL.
- `CM` = Sales Value - Total DL - Total Materials.

## Wynik XLSX

Plik wynikowy zawiera arkusze:
- Profitability
- czasy
- Kliki
- Pivot farby
- zlec + faktury
- faktury linie
- Stawki
- Koszty klików
- Prepress
- Parametry
- Podsumowanie
- Batch summary
- Kokpit

W arkuszu `Profitability` ukrywane są wskazane kolumny techniczne, dodawane są filtry, zamrożenie pierwszego wiersza i kolorowanie grup `Total DL` oraz `Total Materials`.
