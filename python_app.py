import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askopenfilename
import tkinter.font as font
import pandas as pd
import os


# Fonction pour afficher une popup d'information
def popupmsg(msg):
    popup = tk.Tk()
    popup.wm_title("Votre attention s'il vous plait !")
    popup.configure(background="#337746")
    label_popup = tk.Label(popup, text=msg, foreground="#F8F8FF", background="#337746")
    label_popup['font'] = f_label
    label_popup.pack(padx=10, pady=10, expand=True, fill='both')
    bouton_popup = tk.Button(popup, text="C'est compris!", command=popup.destroy)
    bouton_popup.configure(foreground="black", background="white")
    bouton_popup['font'] = f_bouton
    bouton_popup.pack(padx=10, pady=10, expand=True, fill='both')
    popup.mainloop()


# Fonction pour traiter les fichiers Excel ou CSV
def traitement():
    chemin_fichier = askopenfilename(filetypes=[('Excel', ('*.xls', '*.xlsx')), ('CSV', '*.csv')])
    if chemin_fichier.endswith('.csv'):
        df = pd.read_csv(chemin_fichier)
    else:
        df = pd.read_excel(chemin_fichier, engine='openpyxl')

    chemin_dossier = os.path.dirname(chemin_fichier)
    if "données_modifiées.xlsx" in os.listdir(chemin_dossier):
        popupmsg(
            "Il y a déjà un fichier nommé 'données_modifiées.xlsx' dans le dossier, veuillez renommer ce fichier pour que le programme puisse en générer un nouveau.")
    else:
        df.to_excel(os.path.join(chemin_dossier, "données_modifiées.xlsx"), index=False)
        popupmsg(
            "Votre fichier Excel est généré : Vous pouvez le trouver dans le même dossier que le fichier initial." + "\n" + "\n" + "A bientôt!" + "\n")


# Création de l'interface graphique
interface = tk.Tk()
interface.title("Programme de modification de fichier Excel")
interface.configure(background="#337746")

# Font settings
f_label = font.Font(family='Times New Roman', size=20)
f_bouton = font.Font(family='Times New Roman', size=15, weight="bold")

# Création d'un conteneur pour tous les widgets
conteneur = tk.Frame(interface, background="#337746")
conteneur.grid(row=0, column=0, sticky="nsew")

# Configuration du grid de l'interface principale
interface.grid_rowconfigure(0, weight=1)
interface.grid_columnconfigure(0, weight=1)

# Configuration du grid dans le conteneur
conteneur.grid_rowconfigure(0, weight=1)  # Ligne pour le label de bienvenue
conteneur.grid_rowconfigure(1, weight=1)  # Ligne pour le premier bouton
conteneur.grid_rowconfigure(2, weight=1)  # Ligne pour le texte explicatif
conteneur.grid_rowconfigure(3, weight=1)  # Ligne pour le bouton quitter
conteneur.grid_columnconfigure(0, weight=1)

# 1er message (label)
label = tk.Label(conteneur, text="\n" + "Bonjour! " + "\n" + "\n" + "Veuillez sélectionner votre fichier :" + "\n",
                 foreground="white", background="#337746")
label['font'] = f_label
label.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")  # Utilisation de grid avec expansion

# 1er bouton
bouton = tk.Button(conteneur, text='Cliquez ici pour charger votre fichier', command=traitement)
bouton.configure(foreground="#000000", background="white")
bouton['font'] = f_bouton
bouton.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

# 2ème bouton
bouton2 = tk.Button(conteneur, text="Quitter", command=interface.destroy)
bouton2.configure(foreground="#000000", background="white")
bouton2['font'] = f_bouton
bouton2.grid(row=3, column=0, padx=10, pady=10, sticky="nsew")

# Configuration du redimensionnement automatique
conteneur.grid_rowconfigure(0, weight=1)  # Première ligne pour le label
conteneur.grid_rowconfigure(1, weight=1)  # Ligne du premier bouton
conteneur.grid_rowconfigure(2, weight=1)  # Ligne du deuxième label
conteneur.grid_rowconfigure(3, weight=1)  # Ligne du bouton Quitter

interface.mainloop()
