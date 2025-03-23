import tkinter as tk
from tkinter import ttk, messagebox

class MainForm:
    def __init__(self, root):
        self.root = root
        self.root.title("Console")  # Titre de la fenêtre
        self.root.geometry("1200x800")  # Taille de la fenêtre

        # Créer un Notebook (MultiPage en VBA)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True)

        # Onglet 1 : Global
        self.tab_global = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_global, text="Global")

        # Onglet 2 : Rappels
        self.tab_rappels = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_rappels, text="Rappels")

        # Ajouter des contrôles à l'onglet Global
        self.create_global_tab()

        # Ajouter des contrôles à l'onglet Rappels
        self.create_rappels_tab()

    def create_global_tab(self):
        # Exemple de contrôle : Label
        label_global = tk.Label(self.tab_global, text="Onglet Global")
        label_global.pack(pady=10)

        # Exemple de contrôle : ComboBox
        self.combo_action_global = ttk.Combobox(self.tab_global, values=["Appel", "Mail", "Proposition"])
        self.combo_action_global.pack(pady=10)

        # Exemple de contrôle : Button
        btn_enregistrer = tk.Button(self.tab_global, text="Enregistrer", command=self.btn_enregistrer_click)
        btn_enregistrer.pack(pady=10)

    def create_rappels_tab(self):
        # Exemple de contrôle : ListBox
        self.listbox_rappels = tk.Listbox(self.tab_rappels)
        self.listbox_rappels.pack(fill="both", expand=True, padx=10, pady=10)

        # Exemple de contrôle : Button
        btn_quitter = tk.Button(self.tab_rappels, text="Quitter", command=self.btn_quitter_click)
        btn_quitter.pack(pady=10)

    def btn_enregistrer_click(self):
        # Exemple de gestion d'événement
        selected_action = self.combo_action_global.get()
        messagebox.showinfo("Info", f"Action sélectionnée : {selected_action}")

    def btn_quitter_click(self):
        # Fermer l'application
        self.root.quit()


if __name__ == "__main__":
    root = tk.Tk()
    app = MainForm(root)
    root.mainloop()
