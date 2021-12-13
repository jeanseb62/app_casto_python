from tkinter import *
from tkinter.filedialog import askopenfilename
import tkinter as tk
import tabula
import pandas as pd
import os
import PyPDF2
import re
import openpyxl

chemin = "C:/wamp64/www/app_casto/"

liste_fichiers = os.listdir(chemin)

prix = pd.read_excel(chemin+'/prix/réfs_prix.xlsx',sheet_name ='Feuil1')

for fichier in liste_fichiers:
    #on ne traite que les pdf
    if fichier.split(".")[-1]=="pdf":
        nom = fichier[:-4]

# Création de l'interface graphique
root = tk.Tk()
# Géométrie
root.geometry('400x200')

root.wm_iconbitmap('C:/wamp64/www/app_casto/assets/img/logo.ico')
root.wm_title('RIOU Solutions')

# Sélectionner un fichier PDF
def select_file():
    global file_name
    file_name = askopenfilename(filetypes=[("Excel files", ".pdf")])

# PDF TO EXCEL
def pdf_to_excel():
        df_list = tabula.read_pdf(chemin+"/"+nom+".pdf", lattice = True, pages = 'all')

        #le [1::2] permet de ne pas prendre en compte les en-tête de chaque page du pdf de casto
        gros_df = pd.concat(df_list[1::2], ignore_index=True, sort=False)
        
        matiere = []
        impression = []
        
        #ici on corrige le bug de casto qui conduit à ne pas avoir la matière
        #pour chaque ligne du tableau on compare le prix unitaire à celui de nos matières premières pour retrouver la matière correspondante
        
        for index, produit in gros_df.iterrows():
            price = produit["Prix\rUnitaire €"]

            mat = ''
            imp = ''
            for indexe, ligne in prix.iterrows():
                if price == ligne['Impression Recto']:
                    mat = ligne['Détail'] 
                    imp = 'Recto'
                elif price == ligne['Impression Recto/Verso']:
                    mat = ligne['Détail'] 
                    imp = 'Recto-Verso'
                elif price == ligne['Sans impression']:
                    mat = ligne['Détail'] 
                    imp = 'Sans Impression'
            matiere.append(mat)
            impression.append(imp)

        #on ajoute les 2 colonnes matière et impression qui contiennent les infos manquantes
       
        gros_df['surface m2'] = (gros_df['Largeur'] / 1000)*(gros_df['Longueur'] / 1000) *gros_df['Qté']


        gros_df['matiere'] = matiere
        gros_df['impression'] = impression

        

        #on supprime ensuite les colonnes inutiles
        col_suppr = []
        for col in gros_df.columns :
            if 'Unnamed' in col :
                col_suppr.append(col)

        for c in col_suppr:
            gros_df.drop(c, inplace = True, axis = 1)

        #on sauvegarde le tableau dans le fichier excel du nom de notre choix
        
        
        gros_df.replace('',float("NaN"), inplace=True)
        gros_df.dropna(thresh=3,inplace=True)
        gros_df.reset_index(drop = True, inplace = True)
        gros_df.insert(2,'Type', "P")
        gros_df.loc[2,'Type'] = "E"
        gros_df.to_excel(chemin+"/"+nom+".xlsx",index=False)


         #La partie suivante gère l'entête du fichier casto pour récupérer les infos souhaitées (n° de commande, date de livraison, magasin à livrer)

        #On ouvre à nouveau le pdf
    
        pdfFileObj = open(chemin +"/"+ nom+'.pdf', 'rb') 
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj) 
        #on se contente de la première page
        
        pageObj = pdfReader.getPage(0) 
        Texte = pageObj.extractText()
        #on commence après le '.fr' qui est à la fin de la page mais apparait ici au début de la string (je ne sais pas pourquoi). Les infos avant ce '.fr' sont inutiles
        
        Texte = Texte.split('.fr')[-1]

        #On considère que toute majuscule précédée d'une minuscule (et non d'un espace) marque le début d'un nouveau "champ". 
        
        cases = re.split(r"([a-z][A-Z])", Texte)


        #regex_Designation_no = re.compile(r"Désignation(INV-\d+)")
        #Designation = re.search(regex_Designation_no, Texte).group(1)
    
        #Cette boucle permet la bonne séparation des cases et remet la minuscule à la fin de la précédente, la majuscule au début de la suivante
        
        i = 0
        while i < len(cases) :
            if len(cases[i])==2:
                cases[i+1] = cases[i][1] + cases[i+1]
                cases[i-1] = cases[i-1] + cases[i][0]

                cases.pop(i)
                i = i-1
            i+=1
        
        
        #On répète cette opération pour les majuscules précédées d'un chiffre. Attention, si jamais une référence du type "AZER65QDSRR" était dans le fichier, elle serait coupée. 
        #Cependant ça n'était pas le cas jusqu'à présent et à priori cela pourrait ne pas géner le fonctionnement de l'outil suivant la position de cette référence
        
        for i in range(len(cases)):
            cases[i] = re.split(r"([0-9][A-Z])", cases[i])



        cases = [item for sublist in cases for item in sublist]

        j = 0
        while j < len(cases) :
            if len(cases[j])==2:
                cases[j+1] = cases[j][1] + cases[j+1]
                cases[j-1] = cases[j-1] + cases[j][0]

                cases.pop(j)
                j = j-1
            j+=1

    
        pdfFileObj.close()

        #enfin on parcourt toutes les cases pour récupérer les infos que l'on cherche 
        for case in cases:
            if 'Numéro de commande' in case :
                num_com = case.split(':')[-1]
            if 'Magasin :' in case :
                magasin = case.split(':')[-2][:-5]
            if 'Date :' in case :
                date = case.split(':')[-1] 
           
                             
                        

        #dernière partie : on ouvre l'excel du devis et on y ajoute une feuille nommée 'Informations Client' contenant ces informations
        wb = openpyxl.load_workbook(chemin+"/"+nom+".xlsx")

        wb_sheet = wb.active

        #wb_sheet.append(["Désignation"])


        wb_sheet['M1'] = 'Numéro de commande'
        wb_sheet['M2'] = num_com
        wb_sheet['N1'] = 'Magasin'
        wb_sheet['N2'] = magasin
        wb_sheet['O1'] = 'Date'
        wb_sheet['O2'] = date
        

        #data=[num_com,magasin,Type,Designation,ref_Frn,ref_dtm,Qté,surface,matiere,impression,Longueur, Largeur,Total,Prix]
        #df = pd.DataFrame(data)
       # print(df) 


        wb.save(chemin+"/"+nom+".xlsx")
        wb.close()



# PDF TO CSV
'''def pdf_to_csv():
    if file_name.endswith('.pdf'):
        # Read PDF File
        # this contain a list
        df = tabula.read_pdf(file_name, pages = 1)[0]

        # Convert into CSV File
        df.to_csv(chemin+"/"+nom+".csv",index=False)'''


# Add Labels and Buttons
Label(root, text="Convertir PDF en EXCEL", font="italic 15 bold").pack(pady=10)

Button(root,text="Selectionner le fichier PDF",command=select_file,font=14).pack(pady=10)

frame = Frame()
frame.pack(pady=20)

excel_btn = Button(frame,text="PDF en Excel",command=pdf_to_excel,relief="solid",
                   bg="white",font=15)
excel_btn.pack(side=LEFT,padx=10)

#csv_btn = Button(frame,text="PDF en CSV",command=pdf_to_csv,relief="solid",
               #  bg="white",font=15)
#csv_btn.pack()

# Execute Tkinter
root.mainloop()