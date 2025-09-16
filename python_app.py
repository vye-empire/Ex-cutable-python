#Importation des différentes bibliothèque nécessaire au bon fonctionnement du code
import tkinter as tk
from tkinter import messagebox, ttk
from tkinter.filedialog import*
import csv
import tkinter.font as tkfont

#Création d'une classe de données géographique
class Donnee:
    """
    Classe représentant une donnée géographique.
    """
    def __init__(self, id_appuie, num_appuie, code_centre, periode):
        self.id_appuie = id_appuie
        self.num_appuie = num_appuie
        self.code_centre = code_centre
        self.periode = periode

# Fonction pour transformer les données du fichier CSV en objets Donnee
def liste_transformee(liste):
    lis = []
    for i in range(len(liste)):
        try:
            lis.append(Donnee(liste[i][8].strip(), liste[i][9].strip(), liste[i][0].strip(), liste[i][3].strip()))
        except IndexError:
            print(f"Ligne ignorée à cause de l'indice incorrect")
    return lis

# Fonction pour lire un fichier CSV
def lecture_fichier_csv(nom_fichier, delimitateur=';', encodage='ISO-8859-1', nb_lignes_entete=1):
    try:
        with open(nom_fichier, 'r', encoding=encodage) as fichier:
            cr = csv.reader(fichier, delimiter=delimitateur)
            for _ in range(nb_lignes_entete):
                next(cr)
            resultat = [ligne for ligne in cr]
        return resultat
    except Exception as e:
        print(f"Erreur lors de la lecture du fichier : {e}")
        return []

# Fonction pour afficher les résultats sur l'interface graphique
def afficher_donnees(liste_donnees, text_widget):
    text_widget.delete('1.0', tk.END)  # Vider le widget de texte
    for donnee in liste_donnees:
        text_widget.insert(tk.END, f"ID_Appuie : {donnee.id_appuie},     Numéro d'Appuie : {donnee.num_appuie},      "
                                   f"Code Centre : {donnee.code_centre},     Période : {donnee.periode},\n")

# Fonction pour rechercher des données selon un critère
def recherche_donnee(liste_donnees, critere, valeur):
    valeur = valeur.strip()  # Nettoyer la valeur recherchée
    print(f"Recherche par critère : {critere}, valeur : {valeur}")  # Debugging print

    if critere == "ID_Appuie":
        resultats = [d for d in liste_donnees if d.id_appuie == valeur]
    elif critere == "Numéro d'Appuie":
        resultats = [d for d in liste_donnees if d.num_appuie == valeur]
    elif critere == "Code centre":
        resultats = [d for d in liste_donnees if d.code_centre == valeur]
    elif critere == "Période":
        resultats = [d for d in liste_donnees if d.periode == valeur]
    else:
        messagebox.showerror("Erreur", "Critère de recherche invalide")
        return []

    print(f"Nombre de résultats trouvés : {len(resultats)}")  # Debugging print
    messagebox.showinfo("Résultats", f"{len(resultats)} résultat(s) trouvé(s).")

    return resultats

# Fonction pour demander à l'utilisateur s'il veut afficher les résultats
def demander_affichage_resultats(resultats, text_widget):
    if not resultats:
        return

    # Demander à l'utilisateur s'il veut afficher les résultats
    reponse = messagebox.askyesno("Afficher les résultats", "Voulez-vous afficher les résultats de la recherche ?")

    if reponse:  # Si l'utilisateur a répondu oui
        afficher_donnees(resultats, text_widget)

# Fonction pour charger un fichier et afficher les données
def charger_fichier(text_widget):
    global liste_donnees
    chemin_fichier = askopenfilename(filetypes=[('CSV files', '*.csv')])
    if not chemin_fichier:
        return

    popup_traitement = tk.Toplevel()
    popup_traitement.title("Traitement des données")
    label_traitement = tk.Label(popup_traitement, text="Les données sont en cours de traitement...")
    label_traitement.pack(padx=20, pady=20)

    # Utiliser la méthode after pour simuler un délai et fermer le popup après traitement
    def traitement():
        fichier = lecture_fichier_csv(chemin_fichier, delimitateur=";", encodage="ISO-8859-1", nb_lignes_entete=1)
        global liste_donnees
        liste_donnees = liste_transformee(fichier)
        afficher_donnees(liste_donnees, text_widget)
        popup_traitement.destroy()

    text_widget.after(100, traitement)


def charger_donnees_traitees(text_widget):
    global liste_donnees
    chemin_fichier = askopenfilename(filetypes=[('CSV files', '*.csv')])
    if not chemin_fichier:
        return

    popup_traitement = tk.Toplevel()
    popup_traitement.title("Traitement des données")
    label_traitement = tk.Label(popup_traitement, text="Les données sont en cours de traitement...")
    label_traitement.pack(padx=(90,0), pady=20)

    try:
        # Lecture du fichier et affichage direct dans le widget
        with open(chemin_fichier, 'r', encoding='ISO-8859-1') as fichier:
            cr = csv.reader(fichier)
            text_widget.delete('1.0', tk.END)  # Vider le widget avant d'afficher le nouveau contenu
            for ligne in cr:
                text_widget.insert(tk.END, '    '.join(ligne) + '\n')
        messagebox.showinfo("Chargement réussi", "Le fichier a été chargé et affiché avec succès.")
    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur est survenue lors du chargement du fichier : {e}")
    text_widget.after(100, charger_donnees_traitees)

def exporter_donnees(liste_donnees):
    # Demande le nom du fichier et l'emplacement
    chemin_fichier = asksaveasfilename(defaultextension=".csv",filetypes=[('CSV files', '*.csv')])
    if not chemin_fichier:
        return

    try:
        with open(chemin_fichier, 'w', newline='', encoding='UTF-8') as fichier:
            writer = csv.writer(fichier)
            # Écrire l'en-tête
            writer.writerow(['ID_Appuie', 'Numéro d\'Appuie', 'Code Centre', 'Période'])
            # Écrire les données
            for donnee in liste_donnees:
                writer.writerow([donnee.id_appuie, donnee.num_appuie, donnee.code_centre, donnee.periode])
        messagebox.showinfo("Exportation réussie", "Les données ont été exportées avec succès.")
    except Exception as e:
        messagebox.showerror("Erreur d'exportation", f"Une erreur est survenue lors de l'exportation : {e}")


# Fonction pour réinitialiser les critères de recherche et afficher toutes les données
def reinitialiser_recherche(text_widget, critere_var, valeur_entry):
    valeur_entry.delete(0, tk.END)
    afficher_donnees(liste_donnees, text_widget)

# Interface graphique avec Tkinter
def creer_interface():
    global liste_donnees
    liste_donnees = []

    root = tk.Tk()
    root.title("Manipulation de données géographiques")
    # Définir la taille initiale de la fenêtre
    #root.geometry("1857x1011")
    root.resizable(False,False)
    root.attributes("-fullscreen",True)

    # Zone de texte pour afficher les données
    font = tkfont.Font(size=14)
    text_widget = tk.Text(root, height=10, width=110, font=font)
    text_widget.grid(row=0, column=0, columnspan=4, padx=10, pady=(30,0), sticky="nsew")

    # Bouton pour charger un fichier
    font = tkfont.Font(size=14)
    bouton_charger = tk.Button(root, text="Traiter un fichier",fg="white", bg="#337746", command=lambda: charger_fichier(text_widget),font=font)
    bouton_charger.grid(row=1, column=0, padx=(90,10), pady=20, sticky="")

     # Bouton pour charger un fichier déja traité
    font = tkfont.Font(size=14)
    bouton_chargerdonne = tk.Button(root, text="Charger un fichier",fg="white", bg="#337746", command=lambda:charger_donnees_traitees(text_widget),font=font)
    bouton_chargerdonne.grid(row=2, column=0, padx=(100,10), pady=20, sticky="")




    # Champ de saisie et bouton pour la recherche
    font=tkfont.Font(size=14)
    tk.Label(root, text="Critère :",font=font).grid(row=1, column=1, padx=(10,1300), pady=5)
    critere_var = tk.StringVar(value="Sélectionner un critère")
    critere_entry = ttk.Combobox(root, textvariable=critere_var,font=font)
    critere_entry['values'] = ("ID_Appuie", "Numéro d'Appuie", "Code centre", "Période")
    critere_entry.grid(row=1, column=1, padx=(10,960), pady=5)
    tk.Label(root, text="Valeur :",font=font).grid(row=2, column=1, padx=(10,1300), pady=5)
    valeur_entry = tk.Entry(root,font=font)
    valeur_entry.grid(row=2, column=1, padx=(10,978), pady=5)
    

    def lancer_recherche():
        critere = critere_var.get()
        valeur = valeur_entry.get()
        print(f"Lancer recherche avec critère : {critere}, valeur : {valeur}")  # Debugging print
        resultats = recherche_donnee(liste_donnees, critere, valeur)
        demander_affichage_resultats(resultats, text_widget)

    bouton_recherche = tk.Button(root, text="Rechercher", bg="#337746", command=lancer_recherche,font=font,fg="white")
    bouton_recherche.grid(row=2, column=1, columnspan=1, padx=(10,600),pady=0)

    # Bouton pour réinitialiser
    bouton_reinitialiser = tk.Button(root, text="Réinitialiser la recherche", bg="#337746",
                                     command=lambda: reinitialiser_recherche(text_widget, critere_var, valeur_entry),font=font,fg="white")
    bouton_reinitialiser.grid(row=2, column=1, columnspan=1, padx=(10,100))

    # Bouton pour exporter les données
    bouton_exporter = tk.Button(root, text="Exporter les données", bg="#337746",fg="white",font=font, command=lambda: exporter_donnees(liste_donnees))
    bouton_exporter.grid(row=3, column=1, padx=(800,200), pady=5)


    # Bouton pour quitter
    bouton_quitter = tk.Button(root, text="Quitter", bg="#337746",fg="white", command=root.quit,font=font)
    bouton_quitter.grid(row=5, column=0, columnspan=4, padx=(77,0),pady=(0,30))

    # Configurer les lignes et colonnes pour permettre le redimensionnement
    root.grid_rowconfigure(0, weight=1)  # La première rangée (zone de texte) doit s'étendre
    root.grid_columnconfigure(0, weight=1)  # La première colonne doit s'étendre
    root.grid_columnconfigure(1, weight=1)  # La deuxième colonne doit s'étendre

    # Gérer l'événement de redimensionnement de la fenêtre pour ajuster la taille du widget de texte
    text_widget.bind("<Configure>", lambda e: text_widget.config(width=e.width, height=e.height))

    root.mainloop()

# Exécution du programme
if __name__ == "__main__":
    creer_interface()

