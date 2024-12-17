import customtkinter as ctk
from tkinter import Toplevel
from tktimepicker import AnalogPicker, AnalogThemes

# Initialiser CustomTkinter
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class AnalogTimePickerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configuration de la fenêtre principale
        self.title("Sélecteur d'Heure Stylé")
        self.geometry("400x300")

        # Bouton pour ouvrir le sélecteur d'heure
        self.time_button = ctk.CTkButton(self, text="Choisir une heure", command=self.open_time_picker)
        self.time_button.pack(pady=20)

        # Label pour afficher l'heure sélectionnée
        self.time_label = ctk.CTkLabel(self, text="Heure sélectionnée : Aucune", font=("Arial", 16))
        self.time_label.pack(pady=10)

    def open_time_picker(self):
        """Ouvre une fenêtre pop-up avec un sélecteur d'heure analogique."""
        popup = Toplevel(self)
        popup.title("Sélecteur d'Heure")
        popup.geometry("400x400")

        # Création du sélecteur d'heure analogique
        time_picker = AnalogPicker(popup)
        time_picker.pack(expand=True, fill="both", pady=10)

        # Appliquer un thème au sélecteur
        theme = AnalogThemes(time_picker)
        theme.setDracula()  # Exemple : Thème "Dracula"

        # Bouton pour confirmer l'heure
        def confirm_time():
            selected_time = time_picker.time()  # Récupérer l'heure sélectionnée
            self.time_label.configure(text=f"Heure sélectionnée : {selected_time}")
            popup.destroy()

        confirm_button = ctk.CTkButton(popup, text="Confirmer", command=confirm_time)
        confirm_button.pack(pady=20)

# Lancer l'application
if __name__ == "__main__":
    app = AnalogTimePickerApp()
    app.mainloop()
