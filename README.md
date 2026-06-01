Cashbook OCR -> Excel (v5.1)

Jest to program napisany dla księgowości, zajmuje się pobieraniem danych z programu cashbook i zapisywaniem go do formatu xlsx.

Bazuję na działającym v5 i dodaję postprocessing po OCR:
- globalna naprawa numeracji,
- czyszczenie śmieciowych tokenów w kategoriach / opisie,
- normalizacja kwot,
- próba odzyskania brakującej wpłaty/wypłaty na podstawie bilansu,
- usuwanie pustych / duchowych rekordów.

Uruchomienie:
    python cashbook_ocr_to_excel_v5_1.py "snipping tool-20260530T093141Z-3-001.zip" wynik.xlsx
""" 
