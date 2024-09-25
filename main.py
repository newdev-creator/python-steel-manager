import json
import os
from datetime import datetime
from tkinter import messagebox, simpledialog, ttk
from typing import List, Dict, Any, Optional

import customtkinter as ctk
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle

# Path to the JSON file where entries will be saved
FILE_PATH = "data.json"

# List of materials to be used in the dropdown menu
material_list = ["Acier", "Cuivre", "Laiton", "Bronze", "Titane", "Fonte", "Inox 304L", "Inox 316L", "Inox 174PH",
                 "Inox 15-5PH", "Aluminium 7075", "Aluminium 2017"]


# Class representing an entry in the form
class Entry:
    def __init__(self, entry_id: int, date: str, matiere: str, poids: int) -> None:
        self.id: int = entry_id
        self.date: str = date
        self.matiere: str = matiere
        self.poids: int = poids


# Class to handle data loading and saving
class DataHandler:
    @staticmethod
    def load_data() -> List[Entry]:
        """Loads entries from the JSON file."""
        if os.path.exists(FILE_PATH):
            with open(FILE_PATH, "r") as file:
                data = json.load(file)
                return [Entry(e["id"], e["date"], e["matiere"], e["poids"]) for e in data]
        return []

    @staticmethod
    def save_data(entries: List[Entry]) -> None:
        """Saves the current list of entries to the JSON file."""
        data = [{"id": entry.id, "date": entry.date, "matiere": entry.matiere, "poids": entry.poids} for entry in
                entries]
        with open(FILE_PATH, "w") as file:
            json.dump(data, file, indent=4)


# Main class for the form application
class FormulaireApp:
    def __init__(self, root: ctk.CTk) -> None:
        self.root: ctk.CTk = root
        self.root.title("Formulaire")

        self.password: str = "1992"  # Password for restricted actions
        self.entries: List[Entry] = DataHandler.load_data()
        self.id_counter: int = self.entries[-1].id + 1 if self.entries else 1

        self.init_style()  # Initialize the styles for the widgets
        self.create_widgets()  # Create the form widgets
        self.update_treeview()  # Update the Treeview with current data

    def init_style(self) -> None:
        """Initialize styles for the Treeview widget."""
        self.style: ttk.Style = ttk.Style()
        self.style.configure("Treeview", font=("Roboto", 12), background="#2e2e2e",
                             foreground="#e0e0e0", fieldbackground="#2e2e2e", rowheight=20)
        self.style.configure("Treeview.Heading", font=("Roboto", 18, "bold"),
                             background="#1f1f1f", foreground="#ffffff")
        self.style.map("Treeview", background=[("selected", "#1f6aa5")], foreground=[("selected", "#ffffff")])

    def create_widgets(self) -> None:
        """Creates all the form widgets."""
        self.create_labels()
        self.create_entry_fields()
        self.create_buttons()
        self.tree: ttk.Treeview = self.create_treeview()

    def create_labels(self) -> None:
        """Create labels for the form fields."""
        ctk.CTkLabel(self.root, text="Matière", font=("Roboto", 25)).grid(row=0, column=0, padx=20, pady=(10, 20))
        ctk.CTkLabel(self.root, text="Poids", font=("Roboto", 25)).grid(row=1, column=0, padx=20, pady=(10, 20))

    def create_entry_fields(self) -> None:
        """Create entry fields for material and weight."""
        self.matiere_var: ctk.StringVar = ctk.StringVar(value="Inox")
        self.dropdown_matiere: ctk.CTkOptionMenu = ctk.CTkOptionMenu(self.root, variable=self.matiere_var,
                                                                     values=material_list)
        self.dropdown_matiere.grid(row=0, column=1)

        self.txt_poids: ctk.CTkEntry = ctk.CTkEntry(self.root)
        self.txt_poids.grid(row=1, column=1)

    def create_buttons(self) -> None:
        """Create buttons for validation, deletion, and export actions."""
        ctk.CTkButton(self.root, text="Valider", command=self.add_entry, fg_color="green",
                      hover_color="darkgreen").grid(row=2, column=0, columnspan=2, padx=20, pady=(10, 20))

        ctk.CTkButton(self.root, text="Supprimer la ligne selectionnée",
                      command=lambda: self.check_password(self.delete_selected_entry), fg_color="red",
                      hover_color="darkred").grid(row=3, column=0, columnspan=2, padx=20, pady=10)

        ctk.CTkButton(self.root, text="Exporter en PDF", command=lambda: self.check_password(self.export_to_pdf)).grid(
            row=5, column=0, columnspan=2, padx=20, pady=10)

        ctk.CTkButton(self.root, text="Vider le Tableau", command=lambda: self.check_password(self.clear_table),
                      fg_color="red", hover_color="darkred").grid(row=7, column=0, columnspan=2, padx=20, pady=(0, 10))

    def create_treeview(self) -> ttk.Treeview:
        """Create the Treeview table to display entries."""
        tree = ttk.Treeview(self.root, columns=("ID", "Date", "Matière", "Poids"), show="headings")
        tree.heading("ID", text="ID")
        tree.heading("Date", text="Date")
        tree.heading("Matière", text="Matière")
        tree.heading("Poids", text="Poids")
        tree.grid(row=4, column=0, columnspan=2, padx=20, pady=10, sticky="nsew")
        return tree

    def check_password(self, action: Any) -> None:
        """Prompt for a password before executing restricted actions."""
        entered_password: Optional[str] = simpledialog.askstring("Mot de passe", "Entrez le mot de passe:", show='*')
        if entered_password == self.password:
            action()
        else:
            messagebox.showerror("Erreur", "Mot de passe incorrect.")

    def add_entry(self) -> None:
        """Add a new entry to the list after validation and weight adjustment."""
        matiere: str = self.matiere_var.get()
        poids: str = self.txt_poids.get()

        if not self.validate_entry(poids):
            return

        # Subtract 45 kg from the entered weight
        modif_weight: int = int(poids) - 45

        entry = Entry(self.id_counter, datetime.now().strftime("%d/%m/%Y"), matiere, modif_weight)
        self.entries.append(entry)
        self.id_counter += 1
        DataHandler.save_data(self.entries)
        self.update_treeview()
        self.txt_poids.delete(0, ctk.END)

    def validate_entry(self, poids: str) -> bool:
        """Validate that the weight entry is a valid number."""
        if not poids:
            messagebox.showerror("Erreur", "Veuillez remplir le champ Poids.")
            return False
        try:
            int(poids)
        except ValueError:
            messagebox.showerror("Erreur", "Le poids doit être un nombre.")
            return False
        return True

    def update_treeview(self) -> None:
        """Update the Treeview to reflect current entries."""
        for i in self.tree.get_children():
            self.tree.delete(i)
        for entry in self.entries:
            self.tree.insert("", "end", values=(entry.id, entry.date, entry.matiere, entry.poids))

    def delete_selected_entry(self) -> None:
        """Delete the selected entry from the Treeview and the entries list."""
        selected_item = self.tree.selection()

        if not selected_item:
            messagebox.showerror("Erreur", "Veuillez sélectionner une ligne à supprimer.")
            return

        entry_id = self.tree.item(selected_item, "values")[0]

        # Filter out the entry to be deleted
        self.entries = [entry for entry in self.entries if str(entry.id) != entry_id]
        DataHandler.save_data(self.entries)

        # Remove from Treeview
        self.tree.delete(selected_item)

        messagebox.showinfo("Succès", "L'entrée a été supprimée.")

    def export_to_pdf(self) -> None:
        """Export the current entries to a PDF file."""
        current_date: str = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
        filename: str = self.prepare_pdf_directory(f"tableau_{current_date}.pdf")
        doc = SimpleDocTemplate(filename, pagesize=A4)
        elements = [self.create_pdf_title(), self.create_pdf_date(), self.create_data_table(),
                    self.create_pdf_subtitle(), self.create_weight_summary_table()]
        doc.build(elements)
        messagebox.showinfo("Exportation", f"Le tableau a été exporté en PDF sous le nom {filename}.")

    def prepare_pdf_directory(self, filename: str) -> str:
        """Ensure the PDF directory exists and return the file path."""
        pdf_directory = 'pdf'
        os.makedirs(pdf_directory, exist_ok=True)
        return os.path.join(pdf_directory, filename)

    def create_pdf_title(self) -> Paragraph:
        """Create the PDF title."""
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(name="Centered", parent=styles['Title'], alignment=TA_CENTER, spaceAfter=10)
        return Paragraph("Tableau des Entrées", title_style)

    def create_pdf_date(self) -> Paragraph:
        """Create the PDF date."""
        styles = getSampleStyleSheet()
        normal_style = ParagraphStyle(name="Centered", parent=styles['Normal'], alignment=TA_CENTER, spaceAfter=10)
        return Paragraph(f"Date de création: {datetime.now().strftime('%d/%m/%Y')}", normal_style)

    def create_data_table(self) -> Table:
        """Create the data table for the PDF."""
        data = [["ID", "Date", "Matière", "Poids"]] + [[entry.id, entry.date, entry.matiere, entry.poids] for entry in
                                                       self.entries]
        table = Table(data)
        self.style_table(table)
        return table

    def style_table(self, table: Table) -> None:
        """Apply styling to the PDF table."""
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ]))

    def create_pdf_subtitle(self) -> Paragraph:
        """Create the subtitle for the PDF weight summary section."""
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(name="Centered", parent=styles['Title'], alignment=TA_CENTER, spaceBefore=10,
                                     spaceAfter=10)
        return Paragraph("Tableau des sommes", title_style)

    def create_weight_summary_table(self) -> Table:
        """Create a table of material weight summaries for the PDF."""
        matiere_dict = self.calculate_weight_summary()
        sum_data = [["Matière", "Somme des Poids"]] + [[matiere, poids] for matiere, poids in matiere_dict.items()]
        sum_table = Table(sum_data)
        self.style_table(sum_table)
        return sum_table

    def calculate_weight_summary(self) -> Dict[str, int]:
        """Calculate the total weight for each material."""
        matiere_dict: Dict[str, int] = {}
        for entry in self.entries:
            if entry.matiere in matiere_dict:
                matiere_dict[entry.matiere] += entry.poids
            else:
                matiere_dict[entry.matiere] = entry.poids
        return matiere_dict

    def clear_table(self) -> None:
        """Clear the entire table and reset entries."""
        if messagebox.askokcancel("Confirmation",
                                  "Êtes-vous sûr de vouloir vider le tableau ? Cette action est irréversible."):
            self.entries.clear()
            self.id_counter = 1
            DataHandler.save_data(self.entries)
            self.update_treeview()
            messagebox.showinfo("Vider", "Le tableau a été vidé.")


# Main entry point for the application
if __name__ == "__main__":
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()  # Create the root window
    app = FormulaireApp(root)  # Instantiate the form application
    root.mainloop()  # Run the main loop
