import os
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox, Toplevel
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from datetime import datetime, timedelta
#useless.py
# Initialiser l'application
ctk.set_appearance_mode("System")  # Options : "Light", "Dark", "System"
ctk.set_default_color_theme("blue")  # Options : "blue", "green", "dark-blue"

class Application(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Traitement Excel avec Calculs de Temps")
        self.geometry("800x600")

        # Variables
        self.fichier_excel = None
        self.production_par_heure = None
        self.date_heure_debut = None
        self.prefixe_fichier = None

        # Interface graphique
        self.label_instruction = ctk.CTkLabel(self, text="Logiciel de traitement", font=("Arial", 16))
        self.label_instruction.pack(pady=10)

        self.button_choisir_fichier = ctk.CTkButton(self, width=250, text="Choisir un fichier Excel (.xls .xlsx .xlsm et .xlsb)", command=self.demander_fichier)
        self.button_choisir_fichier.pack(pady=10)

        self.label_production = ctk.CTkLabel(self, text="Production par heure (UVC/h):", font=("Arial", 14))
        self.label_production.pack(pady=5)

        self.entry_production = ctk.CTkEntry(self, placeholder_text="Entrez une valeur", width=250)
        self.entry_production.pack(pady=5)

        self.label_date_heure = ctk.CTkLabel(self, text="Date/Heure de début (YYYY-MM-DD HH:MM:SS):", font=("Arial", 14))
        self.label_date_heure.pack(pady=5)

        self.entry_date_heure = ctk.CTkEntry(self, placeholder_text="Entrez la date et l'heure", width=250)
        self.entry_date_heure.pack(pady=5)

        # Nouveau champ pour le préfixe
        self.label_prefixe = ctk.CTkLabel(self, text="Préfixe pour le fichier final :", font=("Arial", 14))
        self.label_prefixe.pack(pady=5)

        self.entry_prefixe = ctk.CTkEntry(self, placeholder_text="Ex : Résultat_", width=250)
        self.entry_prefixe.pack(pady=5)

        self.button_lancer = ctk.CTkButton(self, width=250, text="Lancer le traitement", command=self.lancer_traitement)
        self.button_lancer.pack(pady=20)

        self.button_preview = ctk.CTkButton(self, width=250, text="Aperçu du fichier", command=self.preview_fichier)
        self.button_preview.pack(pady=10)

        self.label_copyright = ctk.CTkLabel(self, text="© Yanis Bordonado - Tous droits réservés", font=("Arial", 12), anchor="e")
        self.label_copyright.pack(side="bottom", anchor="se", pady=10, padx=10)

    def demander_fichier(self):
        self.fichier_excel = filedialog.askopenfilename(
            title="Sélectionner le fichier Excel",
            filetypes=[("Fichiers Excel", "*.xls;*.xlsx;*.xlsm;*.xlsb")]
        )
        if self.fichier_excel:
            messagebox.showinfo("Succès", f"Fichier sélectionné : {self.fichier_excel}")
        else:
            messagebox.showerror("Erreur", "Aucun fichier sélectionné.")

    def lancer_traitement(self):
        try:
            # Récupérer les valeurs de l'interface
            self.production_par_heure = float(self.entry_production.get())
            self.date_heure_debut = datetime.strptime(self.entry_date_heure.get(), "%Y-%m-%d %H:%M:%S")
            self.prefixe_fichier = self.entry_prefixe.get().strip()

            if not self.fichier_excel:
                raise ValueError("Aucun fichier Excel sélectionné.")

            # Charger et traiter le fichier Excel
            self.traiter_fichier_excel()

            messagebox.showinfo("Succès", "Le traitement du fichier a été effectué avec succès.")

        except ValueError as ve:
            messagebox.showerror("Erreur", f"Entrée invalide : {ve}")
        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur s'est produite : {e}")

    def traiter_fichier_excel(self):
        """
        Charge et traite le fichier Excel, en sélectionnant automatiquement le moteur en fonction de l'extension.
        """
        # Vérifier l'extension du fichier
        _, extension = os.path.splitext(self.fichier_excel)

        # Sélectionner le moteur en fonction de l'extension
        if extension == ".xlsx" or extension == ".xlsm":
            moteur = "openpyxl"
        elif extension == ".xls":
            moteur = "xlrd"
        elif extension == ".xlsb":
            moteur = "pyxlsb"
        else:
            raise ValueError(f"Format de fichier non pris en charge : {extension}")

        # Charger le fichier avec le moteur approprié
        try:
            df = pd.read_excel(self.fichier_excel, engine=moteur)
            print(f"Fichier chargé avec succès en utilisant le moteur '{moteur}'.")
        except Exception as e:
            raise ValueError(f"Erreur lors du chargement du fichier avec le moteur '{moteur}': {e}")

        # Traitement des données
        df_principales = df[df['Déclaration\nFin Prépa ?'].isna()]
        colonnes_a_garder = ['Chargement ', 'Date Chargement', 'Nb Prépa', 'UVC à préparer', 'UVC Prépa Validées']
        df_principales = df_principales[colonnes_a_garder]

        # Calculer UVC restants et heures nécessaires
        df_principales['UVC Restants'] = df_principales['UVC à préparer'] - df_principales['UVC Prépa Validées']
        df_principales['Heures nécessaires'] = df_principales['UVC Restants'] / self.production_par_heure

        # Générer les heures de fin
        date_heure = self.date_heure_debut
        current_datetime = date_heure

        heures_de_fin = []
        previous_chargement = None

        for i, row in df_principales.iterrows():
            chargement_value = row[0]
            
            if previous_chargement == "Sous-Total":
                heures_restantes = row['Heures nécessaires']
                fin_datetime = date_heure + timedelta(hours=heures_restantes)
                heures_de_fin.append(fin_datetime.strftime("%Y-%m-%d %H:%M:%S"))
                current_datetime = fin_datetime
            else:
                heures_restantes = row['Heures nécessaires']
                fin_datetime = current_datetime + timedelta(hours=heures_restantes)
                heures_de_fin.append(fin_datetime.strftime("%Y-%m-%d %H:%M:%S"))
                current_datetime = fin_datetime
            
            previous_chargement = chargement_value

        df_principales['Heures de fin'] = heures_de_fin

        # Sauvegarder le fichier traité
        fichier_temporaire = "temp_file.xlsx"
        df_principales.to_excel(fichier_temporaire, index=False)

        # Appliquer la mise en forme conditionnelle
        wb = load_workbook(fichier_temporaire)
        sheet = wb.active

        blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
            chargement_value = row[0].value
            date_chargement = row[1].value
            heure_de_fin = row[-1].value

            if chargement_value == "Sous-Total":
                for cell in row:
                    cell.fill = blue_fill
            elif heure_de_fin and date_chargement:
                heure_de_fin_dt = datetime.strptime(heure_de_fin, "%Y-%m-%d %H:%M:%S")
                date_chargement_dt = datetime.strptime(str(date_chargement), "%Y-%m-%d %H:%M:%S")
                
                if heure_de_fin_dt > date_chargement_dt:
                    for cell in row:
                        cell.fill = red_fill
                elif 0 <= (date_chargement_dt - heure_de_fin_dt).total_seconds() <= 1800:
                    for cell in row:
                        cell.fill = yellow_fill
        self.df_principales = df_principales

        # Ajouter le préfixe au fichier final
        prefixe = self.prefixe_fichier if self.prefixe_fichier else "Fichier_Traité"
        fichier_final = f"{prefixe}.xlsx"
        wb.save(fichier_final)
        os.remove(fichier_temporaire)
    
    def preview_fichier(self):
        """
        Ouvre une nouvelle fenêtre pour afficher un aperçu des données du fichier Excel avec une grille et des couleurs.
        Permet le défilement vertical si les données dépassent la hauteur de la fenêtre.
        """
        if not hasattr(self, 'df_principales') or self.df_principales is None:
            messagebox.showerror("Erreur", "Aucune donnée disponible pour l'aperçu.")
            return

        # Créer une nouvelle fenêtre pour l'aperçu
        preview_window = Toplevel(self)
        preview_window.title("Aperçu des données")
        preview_window.geometry("800x600")  # Ajustez cette taille selon votre besoin

        # Conteneur principal avec barre de défilement
        container = ctk.CTkFrame(preview_window)
        container.pack(fill="both", expand=True)

        canvas = ctk.CTkCanvas(container)
        canvas.pack(side="left", fill="both", expand=True)

        scrollbar = ctk.CTkScrollbar(container, orientation="vertical", command=canvas.yview)
        scrollbar.pack(side="right", fill="y")

        canvas.configure(yscrollcommand=scrollbar.set)

        # Frame intérieure pour afficher le contenu
        inner_frame = ctk.CTkFrame(canvas)
        canvas.create_window((canvas.winfo_width() // 2, canvas.winfo_height() // 2), window=inner_frame, anchor="center")


        # Créer les en-têtes de colonnes avec bordures
        for i, column in enumerate(self.df_principales.columns):
            header_frame = ctk.CTkFrame(inner_frame, border_width=1, border_color="black")  # Ajout de bordures
            header_frame.grid(row=0, column=i, sticky="nsew", padx=1, pady=1)  # Grille serrée avec espace
            label = ctk.CTkLabel(header_frame, text=column, font=("Arial", 12, "bold"), fg_color="lightblue")
            label.pack(fill="both", expand=True)

        # Définir les couleurs selon les critères
        def get_cell_color(row):
            """
            Retourne une couleur en fonction des critères spécifiques à la ligne.
            """
            chargement_value = row[0]
            heure_de_fin = row['Heures de fin']
            date_chargement = row['Date Chargement']

            # Vérifier si les valeurs sont valides avant de les utiliser
            if pd.isna(heure_de_fin) or pd.isna(date_chargement):
                return "white"  # Blanc pour les cellules non valides

            try:
                heure_de_fin_dt = datetime.strptime(heure_de_fin, "%Y-%m-%d %H:%M:%S")
                date_chargement_dt = datetime.strptime(str(date_chargement), "%Y-%m-%d %H:%M:%S")
            except ValueError:
                return "white"  # Blanc pour les formats non valides

            if chargement_value == "Sous-Total":
                return "lightblue"
            elif heure_de_fin_dt > date_chargement_dt:
                return "red"
            elif 0 <= (date_chargement_dt - heure_de_fin_dt).total_seconds() <= 1800:
                return "yellow"
            return "white"


        # Créer les lignes de données avec bordures et couleurs
        for i, row in self.df_principales.iterrows():
            for j, value in enumerate(row):
                cell_color = get_cell_color(row)  # Appliquer la couleur à chaque cellule
                cell_frame = ctk.CTkFrame(inner_frame, border_width=1, border_color="gray", fg_color=cell_color)
                cell_frame.grid(row=i + 1, column=j, sticky="nsew", padx=1, pady=1)
                label = ctk.CTkLabel(cell_frame, text=str(value), font=("Arial", 12))
                label.pack(fill="both", expand=True)

        # Ajuster les colonnes pour qu'elles soient uniformes
        for col in range(len(self.df_principales.columns)):
            inner_frame.grid_columnconfigure(col, weight=1)

        # Centrer le tableau dans la fenêtre
        # Configuration de la grille intérieure pour occuper tout l'espace
        inner_frame.grid_rowconfigure(0, weight=1)  # Premier row a du poids
        inner_frame.grid_columnconfigure(0, weight=1)  # Premier column a du poids

        # Ajouter un padding autour du tableau pour centrer son contenu
        for row in range(1, len(self.df_principales) + 1):
            inner_frame.grid_rowconfigure(row, weight=1)
        for col in range(len(self.df_principales.columns)):
            inner_frame.grid_columnconfigure(col, weight=1)

        # Mettre à jour la taille du canevas pour ajuster au contenu
        inner_frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))

        # Activer le défilement avec la molette
        def on_mousewheel(event):
            canvas.yview_scroll(-1 * int(event.delta / 120), "units")

        canvas.bind_all("<MouseWheel>", on_mousewheel)

        # Bouton pour fermer la fenêtre
        button_close = ctk.CTkButton(preview_window, text="Retour", command=preview_window.destroy)
        button_close.pack(pady=10)








# Lancer l'application
if __name__ == "__main__":
    app = Application()
    app.mainloop()
