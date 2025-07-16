import os
from tkinter import *
from tkinter import ttk, messagebox
from datetime import datetime
from docxtpl import DocxTemplate
from openpyxl import load_workbook, Workbook
from tkinter import Checkbutton

# === Ścieżki do plików ===
TEMPLATE_KARTA = "Karta_basic.docx"
TEMPLATE_ANKIETA = "2_Ankieta potrzeb i plan wsparcia.docx"
EXCEL_PATH = "rejestr konsultacji _basic.xlsx"
OUTPUT_DIR = "output"

if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

# === Aplikacja ===
class App(Tk):
    def __init__(self):
        super().__init__()
        self.title("System ankiet ‑ Mazowieckie Mosty Międzykulturowe")
        self.geometry("1200x780")
        self.resizable(True, True)
        self.interviewer_code = None
        self.show_login()

    def show_login(self):
        self.clear()
        ttk.Label(self, text="Podaj kod ankietera:", font=("Arial", 14)).pack(pady=20)
        self.code_entry = ttk.Entry(self, width=30, font=("Arial", 12))
        self.code_entry.pack()
        ttk.Button(self, text="Zaloguj", command=self.login).pack(pady=10)

    def login(self):
        code = self.code_entry.get().strip()
        if not code:
            messagebox.showerror("Błąd", "Kod ankietera nie może być pusty.")
            return
        self.interviewer_code = code
        self.show_form_choice()

    def show_form_choice(self):
        self.clear()
        ttk.Label(self, text=f"Zalogowano jako: {self.interviewer_code}",
                  font=("Arial", 12)).pack(pady=10)

        ttk.Label(self, text="Wybierz formularz:", font=("Arial", 14)).pack(pady=20)
        ttk.Button(self, text="Karta uczestnika", width=30,
                   command=self.show_karta_form).pack(pady=5)
        ttk.Button(self, text="Ankieta potrzeb", width=30,
                   command=self.show_ankieta_form).pack(pady=5)

    def show_karta_form(self):
        self.clear()
        ttk.Label(self, text="Karta uczestnika", font=("Arial", 16, "bold")).pack(pady=10)

        form = ttk.Frame(self)
        form.pack(pady=10)

        self.karta_vars = {
            "imie": StringVar(),
            "nazwisko": StringVar(),
            "data_urodzenia": StringVar(),
            "plec": StringVar(value="K"),
            "grupa_wiekowa": StringVar(value="18-60"),
            "obywatelstwo": StringVar(),
            "telefon": StringVar(),
            "email": StringVar(),
        }

        self._add_entry(form, "Imię", self.karta_vars["imie"])
        self._add_entry(form, "Nazwisko", self.karta_vars["nazwisko"])
        self._add_entry(form, "Data urodzenia (YYYY-MM-DD)", self.karta_vars["data_urodzenia"])

        ttk.Label(form, text="Płeć:").pack(anchor=W)
        sex_frame = ttk.Frame(form); sex_frame.pack(anchor=W)
        ttk.Radiobutton(sex_frame, text="Kobieta", variable=self.karta_vars["plec"], value="K").pack(side=LEFT)
        ttk.Radiobutton(sex_frame, text="Mężczyzna", variable=self.karta_vars["plec"], value="M").pack(side=LEFT)

        ttk.Label(form, text="Grupa wiekowa:").pack(anchor=W, pady=(5, 0))
        ttk.Combobox(form, textvariable=self.karta_vars["grupa_wiekowa"],
                     values=["poniżej 18", "18-60", "powyżej 60"],
                     state="readonly").pack(anchor=W)

        self._add_entry(form, "Obywatelstwo", self.karta_vars["obywatelstwo"])
        self._add_entry(form, "Telefon", self.karta_vars["telefon"])
        self._add_entry(form, "E‑mail", self.karta_vars["email"])

        ttk.Button(self, text="Zapisz", command=self.save_karta).pack(pady=15)
        ttk.Button(self, text="↩ Wróć", command=self.show_form_choice).pack()

    def save_karta(self):
        data = {k: v.get().strip() for k, v in self.karta_vars.items()}
        if not data["imie"] or not data["nazwisko"]:
            messagebox.showerror("Błąd", "Imię i nazwisko są wymagane.")
            return

        try:
            doc = DocxTemplate(TEMPLATE_KARTA)
            context = {
                "imie": data["imie"],
                "nazwisko": data["nazwisko"],
                "data_urodzenia": data["data_urodzenia"],
                "plec_k": "☒" if data["plec"] == "K" else "☐",
                "plec_m": "☒" if data["plec"] == "M" else "☐",
                "grupa_pon18": "☒" if data["grupa_wiekowa"] == "poniżej 18" else "☐",
                "grupa_18_60": "☒" if data["grupa_wiekowa"] == "18-60" else "☐",
                "grupa_pow60": "☒" if data["grupa_wiekowa"] == "powyżej 60" else "☐",
                "obywatelstwo": data["obywatelstwo"],
                "telefon": data["telefon"],
                "email": data["email"],
                "data_wypelnienia": datetime.today().strftime("%d.%m.%Y"),
                "kod_ankietera": self.interviewer_code,
            }
            # Rozbicie daty urodzenia
            try:
            yyyy, mm, dd = data["data_urodzenia"].split("-")
            except ValueError:
            yyyy, mm, dd = "", "", ""

            context.update(self.rozbij_na_znaki(dd, "d", 2))   # d0, d1
            context.update(self.rozbij_na_znaki(mm, "m", 2))   # m0, m1
            context.update(self.rozbij_na_znaki(yyyy, "r", 4)) # r0, r1, r2, r3

            # Rozbicie telefonu
            context.update(self.rozbij_na_znaki(data["telefon"], "t", 12))  # t0–t11

            filename = f"{data['nazwisko']}_{data['imie']}_{datetime.now():%Y%m%d_%H%M%S}.docx"
            output_path = os.path.join(OUTPUT_DIR, filename)
            doc.render(context)
            doc.save(output_path)
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się wygenerować DOCX:\n{e}")
            return

        print("🛈 Zapis do Excela pominięty – dotyczy tylko Ankiety potrzeb.")
        messagebox.showinfo("Sukces", f"Dane zapisane.\nWygenerowany plik: {output_path}")
        self.show_form_choice()

    def show_ankieta_form(self):
        self.clear()
        ttk.Label(self, text="Ankieta potrzeb", font=("Arial", 16, "bold")).pack(pady=10)

        form = ttk.Frame(self)
        form.pack(pady=10)

        self.ankieta_vars = {
            "imie": StringVar(),
            "nazwisko": StringVar(),
            "data": StringVar(value=datetime.today().strftime("%Y-%m-%d")),
            "inne": StringVar(),
        }

        self.wsparcie_check = {}
        obszary = [
            "wiedza na temat przysługujących uprawnień i możliwości uzyskania pomocy",
            "sprawy urzędowe",
            "legalizacja pobytu",
            "kontakty z instytucjami publicznymi (pomoc społeczna, urzędy, szkoły, służba zdrowia itp.)",
            "nauka języka polskiego",
            "poprawa sytuacji mieszkaniowej",
            "poprawa sytuacji zdrowotnej",
            "dostęp do wsparcia specjalistycznego, np. psycholog, prawnik",
            "kwestie związane z uzyskaniem lub uzupełnieniem wykształcenia",
            "edukacja dzieci (np. rejestracja do przedszkola, szkoły itp.)",
            "adaptacja dzieci w szkole (np. kontakty ze szkołami)",
            "znajomość zasad obowiązujących w polskim społeczeństwie",
        ]

        self._add_entry(form, "Imię", self.ankieta_vars["imie"])
        self._add_entry(form, "Nazwisko", self.ankieta_vars["nazwisko"])
        self._add_entry(form, "Data", self.ankieta_vars["data"])

        ttk.Label(form, text="Obszary wymagające wsparcia:", font=("Arial", 12, "bold")).pack(anchor=W, pady=(10, 0))
        for obs in obszary:
            var = BooleanVar()
            self.wsparcie_check[obs] = var
            Checkbutton(form, text=obs, variable=var, wraplength=600, justify=LEFT).pack(anchor=W)

        self._add_entry(form, "Inne:", self.ankieta_vars["inne"])

        ttk.Button(self, text="Zapisz", command=self.save_ankieta).pack(pady=10)
        ttk.Button(self, text="↩ Wróć", command=self.show_form_choice).pack()

    def save_ankieta(self):
        print("DEBUG: Start zapisu ankiety")

        imie = self.ankieta_vars["imie"].get().strip()
        nazwisko = self.ankieta_vars["nazwisko"].get().strip()
        data = self.ankieta_vars["data"].get().strip()
        inne = self.ankieta_vars["inne"].get().strip()

        if not imie or not nazwisko:
            messagebox.showerror("Błąd", "Imię i nazwisko są wymagane.")
            return

        context = {
            "imie": imie,
            "nazwisko": nazwisko,
            "data": data,
            "inne": inne,
        }
        for obs, var in self.wsparcie_check.items():
            key = self._obszar_to_key(obs)
            context[key] = "☒" if var.get() else "☐"

        try:
            doc = DocxTemplate(TEMPLATE_ANKIETA)
            filename = f"ankieta_{nazwisko}_{imie}_{datetime.now():%Y%m%d_%H%M%S}.docx"
            output_path = os.path.join(OUTPUT_DIR, filename)
            doc.render(context)
            doc.save(output_path)
            print("DEBUG: DOCX zapisany:", output_path)
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się wygenerować DOCX:\n{e}")
            return

        try:
            if not os.path.exists(EXCEL_PATH):
                wb = Workbook()
                ws = wb.active
                ws.append(["NAZWISKO I IMIĘ", "DATA", "OBSZAR WSPARCIA", "LINK DO PLIKU"])
            else:
                wb = load_workbook(EXCEL_PATH)
                ws = wb.active

            ws.append([
                f"{nazwisko} {imie}",
                data,
                self._obszary_lista(),
                output_path,
            ])
            wb.save(EXCEL_PATH)
            print("DEBUG: Dane dopisane do Excela")
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się zapisać danych do Excela:\n{e}")
            return

        messagebox.showinfo("Sukces", f"Ankieta zapisana.\nPlik: {output_path}")
        self.show_form_choice()

    def _obszar_to_key(self, text):
        return "wsparcie_" + text.lower().replace(" ", "_").replace(",", "").replace("–", "-")[:50]

    def _obszary_lista(self):
        zaznaczone = [k for k, v in self.wsparcie_check.items() if v.get()]
        return "; ".join(zaznaczone) if zaznaczone else "(brak)"

    def _add_entry(self, parent, label, variable):
        ttk.Label(parent, text=f"{label}:").pack(anchor=W, pady=(5, 0))
        ttk.Entry(parent, textvariable=variable, width=42).pack(anchor=W)

    def clear(self):
        for widget in self.winfo_children():
            widget.destroy()
            
   def rozbij_na_znaki(self, tekst, prefix, ile_znakow):
       cyfry = ''.join(filter(str.isdigit, tekst))
       return {f"{prefix}{i}": cyfry[i] if i < len(cyfry) else "" for i in range(ile_znakow)}


# === START ===
if __name__ == "__main__":
    app = App()
    app.mainloop()
