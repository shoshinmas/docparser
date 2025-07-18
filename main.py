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
        ttk.Label(self, text=f"Zalogowano jako: {self.interviewer_code}", font=("Arial", 12)).pack(pady=10)
        ttk.Label(self, text="Wybierz formularz:", font=("Arial", 14)).pack(pady=20)
        ttk.Button(self, text="Karta uczestnika", width=30, command=self.show_karta_form).pack(pady=5)
        ttk.Button(self, text="Ankieta potrzeb", width=30, command=self.show_ankieta_form).pack(pady=5)

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
            "miasto": StringVar(),
            "karta_pobytu": StringVar(),
            "wiza": StringVar(),
            "tztc": StringVar(),
            "dokument_podrozy": StringVar(),
            "inny": StringVar(),
            "inny_dokument": StringVar(),
            "dokument": StringVar(),
            "kontakt_niepelnoletni": StringVar(),
            "rodzic": StringVar(),
            "dziadek_babcia": StringVar(),
            "wujek_ciocia": StringVar(),
            "brat_siostra": StringVar(),
            "opiekun_prawny": StringVar(),
            "telefon_opiekuna": StringVar(),
            "email_niepelnoletni": StringVar(),
        }

        self.karta_checkboxy = {
            "status_uczestnika_uchodzca": BooleanVar(),
            "status_uczestnika_ochrona": BooleanVar(),
            "status_uczestnika_pozostaly": BooleanVar(),
            "status_uczestnika_bez": BooleanVar(),
            "status_uczestnika_obywatel": BooleanVar(),
            "status_uczestnika_inny": BooleanVar(),
        }

        for pole in ["imie", "nazwisko", "data_urodzenia", "obywatelstwo", "telefon", "email", "miasto",
                     "karta_pobytu", "wiza", "tztc", "dokument_podrozy", "inny", "inny_dokument", "dokument",
                     "kontakt_niepelnoletni", "rodzic", "dziadek_babcia", "wujek_ciocia", "brat_siostra",
                     "opiekun_prawny", "telefon_opiekuna", "email_niepelnoletni"]:
            self._add_entry(form, pole.replace("_", " ").capitalize(), self.karta_vars[pole])

        ttk.Label(form, text="Płeć:").pack(anchor=W)
        sex_frame = ttk.Frame(form); sex_frame.pack(anchor=W)
        ttk.Radiobutton(sex_frame, text="Kobieta", variable=self.karta_vars["plec"], value="K").pack(side=LEFT)
        ttk.Radiobutton(sex_frame, text="Mężczyzna", variable=self.karta_vars["plec"], value="M").pack(side=LEFT)

        ttk.Label(form, text="Grupa wiekowa:").pack(anchor=W, pady=(5, 0))
        ttk.Combobox(form, textvariable=self.karta_vars["grupa_wiekowa"],
                     values=["poniżej 18", "18-60", "powyżej 60"], state="readonly").pack(anchor=W)

        ttk.Label(form, text="Status uczestnika:").pack(anchor=W, pady=(10, 0))
        for key, var in self.karta_checkboxy.items():
            Checkbutton(form, text=key.replace("status_uczestnika_", "").capitalize(), variable=var).pack(anchor=W)

        ttk.Button(self, text="Zapisz", command=self.save_karta).pack(pady=15)
        ttk.Button(self, text="↩ Wróć", command=self.show_form_choice).pack()

    def save_karta(self):
        data = {k: v.get().strip() for k, v in self.karta_vars.items()}
        if not data["imie"] or not data["nazwisko"]:
            messagebox.showerror("Błąd", "Imię i nazwisko są wymagane.")
            return

        try:
            doc = DocxTemplate(TEMPLATE_KARTA)
            context = data.copy()
            context.update({
                "plec_k": "☒" if data["plec"] == "K" else "☐",
                "plec_m": "☒" if data["plec"] == "M" else "☐",
                "grupa_pon18": "☒" if data["grupa_wiekowa"] == "poniżej 18" else "☐",
                "grupa_18_60": "☒" if data["grupa_wiekowa"] == "18-60" else "☐",
                "grupa_pow60": "☒" if data["grupa_wiekowa"] == "powyżej 60" else "☐",
                "data_wypelnienia": datetime.today().strftime("%d.%m.%Y"),
                "kod_ankietera": self.interviewer_code,
            })

            for key, var in self.karta_checkboxy.items():
                context[key] = "☒" if var.get() else "☐"

            yyyy, mm, dd = (data["data_urodzenia"] + "--").split("-")[:3]
            context.update(self.rozbij_na_znaki(dd, "d", 2))
            context.update(self.rozbij_na_znaki(mm, "m", 2))
            context.update(self.rozbij_na_znaki(yyyy, "r", 4))
            context.update(self.rozbij_na_znaki(data["telefon"], "t", 11))
            context.update(self.rozbij_na_znaki(data["telefon_opiekuna"], "tt", 11))

            filename = f"{data['nazwisko']}_{data['imie']}_{datetime.now():%Y%m%d_%H%M%S}.docx"
            output_path = os.path.join(OUTPUT_DIR, filename)
            doc.render(context)
            doc.save(output_path)
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się wygenerować DOCX:\n{e}")
            return

        messagebox.showinfo("Sukces", f"Dane zapisane.\nWygenerowany plik: {output_path}")
        self.show_form_choice()

    def rozbij_na_znaki(self, tekst, prefix, ile_znakow):
        cyfry = ''.join(filter(str.isdigit, tekst))
        return {f"{prefix}{i}": cyfry[i] if i < len(cyfry) else "" for i in range(ile_znakow)}

    def _add_entry(self, parent, label, variable):
        ttk.Label(parent, text=f"{label}:").pack(anchor=W, pady=(5, 0))
        ttk.Entry(parent, textvariable=variable, width=42).pack(anchor=W)

    def clear(self):
        for widget in self.winfo_children():
            widget.destroy()

if __name__ == "__main__":
    app = App()
    app.mainloop()