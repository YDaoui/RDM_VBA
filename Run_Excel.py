import win32com.client
import os

# Définir le chemin du fichier Excel (ajustez selon votre structure)
repo_path = os.path.dirname(os.path.abspath(__file__))  # Récupérer le chemin du repo
excel_file = os.path.join(repo_path, "Suivie_Batonnage.xlsm")  # Chemin vers le fichier Excel

# Vérifier si le fichier existe
if not os.path.exists(excel_file):
    print(f"Le fichier {excel_file} est introuvable.")
    exit()

# Ouvrir Excel
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True  # Facultatif : rendre Excel visible

# Ouvrir le fichier Excel
wb = excel.Workbooks.Open(excel_file)

# Exécuter la macro du UserForm
try:
    excel.Application.Run("Cons")
    print("UserForm exécuté avec succès.")
except Exception as e:
    print(f"Erreur lors de l'exécution du UserForm : {e}")

# Facultatif : fermer Excel après exécution
# wb.Close(SaveChanges=False)
# excel.Quit()
