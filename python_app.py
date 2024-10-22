import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askopenfilename
import tkinter.font as font
import pandas as pd
import os

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

def traitement():
    chemin_fichier = askopenfilename(filetypes=[('Excel', ('*.xls', '*.xlsx')), ('CSV', '*.csv')])
    if chemin_fichier.endswith('.csv'):
        df = pd.read_csv(chemin_fichier)
    else:
        df = pd.read_excel(chemin_fichier, engine='openpyxl')
    ####################    VOTRE CODE    #####################
    chemin_dossier = os.path.dirname(chemin_fichier)
    if "données_modifiées.xlsx" in os.listdir(chemin_dossier):
        popupmsg("Il y a déjà un fichier nommé 'données_modifiées.xlsx' dans le dossier, veuillez renommer ce fichier pour que le programme puisse en générer un nouveau.")
    else:
        df.to_excel(os.path.join(chemin_dossier, "données_modifiées.xlsx"), index=False)
        popupmsg("Votre fichier excel est généré : Vous pouvez le trouver dans le même dossier que le fichier excel initial." + "\n" + "\n" + "A bientôt!" + "\n")

interface = tk.Tk()
interface.title("Programme de modification de fichier Excel")
interface.configure(background="#337746")

# 1er message
f_label = font.Font(family='Times New Roman', size=20)
f_bouton = font.Font(family='Times New Roman', size=15, weight="bold")
label = tk.Label(interface, text="\n"  + "Bonjour! " + "\n"  + "\n" + "Veuillez sélectionner votre fichier :" + "\n" , foreground="white", background="#337746")
label['font'] = f_label
label.pack(padx=10, pady=10, expand=True, fill='both')

# 1er bouton
bouton = tk.Button(interface, text='Cliquez ici pour charger votre fichier', command=traitement)
bouton.configure(foreground="#000000", background="white")
bouton['font'] = f_bouton
bouton.pack(padx=10, pady=10, expand=True, fill='both')

# 2eme message
label2 = tk.Label(interface, text="\n" + "Une fois le fichier séléctionné, le programme va générer un nouveau fichier excel (nommé 'Données_modifiées.xlsx') dans le même dossier que le fichier séléctionné." + "\n" + "\n" + "Pour relancer le programme, veuillez renommer le fichier 'Données_modifiées.xlsx' pour ne pas rentrer en conflit avec le programme." + "\n", foreground="white", background="#337746")
label2['font'] = f_label
label2.pack(padx=10, pady=10, expand=True, fill='both')

# 2eme bouton
bouton2 = tk.Button(interface, text="Quitter", command=interface.destroy)
bouton2.configure(foreground="#000000", background="white")
bouton2['font'] = f_bouton
bouton2.pack(padx=10, pady=10, expand=True, fill='both')

interface.mainloop()