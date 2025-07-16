# System Ankiet – Mazowieckie Mosty Międzykulturowe

Aplikacja stworzona na potrzeby projektu integracyjnego FAMI **Mazowieckie Mosty Międzykulturowe**.  
Umożliwia ankieterom wygodne wypełnianie formularzy uczestników oraz ankiet potrzeb, z automatycznym generowaniem plików `.docx` i (dla części formularzy) zapisem danych do pliku `.xlsx`.

---

## 🧩 Funkcje

- ✅ Logowanie z użyciem **kodu ankietera**
- ✅ Wybór rodzaju formularza:
  - **Karta uczestnika**
  - **Ankieta potrzeb i plan wsparcia**
- ✅ Graficzny interfejs (Tkinter)
- ✅ Automatyczne generowanie pliku `.docx` na podstawie szablonów Word (`docxtpl`)
- ✅ Zapis danych z ankiety potrzeb do pliku Excel (`openpyxl`)
- ✅ Obsługa pól rozbitych na cyfry (np. data urodzenia, telefon)

---

## 🖥️ Wymagania

- Python 3.8 lub nowszy
- Biblioteki:
  ```bash
  pip install docxtpl openpyxl
