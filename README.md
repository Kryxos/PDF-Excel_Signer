# PDF Podpis — GitHub Pages

Ta wersja działa jako statyczna strona. Nie wymaga EXE ani backendu.

## Wdrożenie

1. Utwórz repozytorium na GitHubie.
2. Wrzuć całą zawartość tego folderu do głównego katalogu repo.
3. Wejdź w **Settings → Pages**.
4. Ustaw **Deploy from a branch**, branch **main**, folder **/(root)**.
5. Zapisz ustawienia.

## Prywatność

PDF, XLSX oraz podpisy są przetwarzane lokalnie w przeglądarce. Aplikacja nie ma backendu i nie wysyła dokumentów do GitHub Pages.

## Biblioteki

- JSZip jest dołączony lokalnie w `vendor/`.
- PDF.js, pdf-lib i ExcelJS są obecnie ładowane z CDN (tak jak w wersji EXE).

## Obsługa XLSX

Patcher XLSX został przeniesiony z backendu Go do JavaScript i działa bezpośrednio w przeglądarce.
