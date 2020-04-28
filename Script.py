# IMPORTATION DES LIBRAIRIES 
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import xlwings as xw 

# POUR NE PAS VOIR LE FICHIER EXCEL S'OUVRIR 
app = xw.App(visible=False)

# IMPORTATION DU FICHIER DE DEPISTAGE
wb = xw.Book("fichier1.xlsm")

# OUVERTURE DU CLASSEUR ET SELECTION DE FEUILLE CONTENANT LE JEU DE DONNEES
sheet = wb.sheets["Feuille_1"]

# In this line, we are initializing "FireFox" by making an object of it.
driver = webdriver.Chrome()
"""The "driver.get method" will navigate to a page given by the URL. WebDriver
 will wait until the page has been completely loaded (that is, the "onload" 
 occasion has let go), before returning control to your test or script."""
driver.get("Lien_de_la_page")


# ATTRIBUTION DES VALEURS DANS LES CHAMPS A RENSEIGNER 
# Les caractères dans la liste "ref" sont communs à l'ensemble des cellules
# Exemple de l'id d'une cellule  "jYMaRS9Bt2Y-mwgEmG5zMCY-val"

ref = ["-mwgEmG5zMCY-val", "-GSpPxY7lzVs-val", "-dcPhrwJ3epH-val", "-GbZK5XGNqbD-val" ]

# Fonctions permettant de réduire le code
def filldata(section_id, section_value,nb = 4):
    for i in range(nb):
        element = driver.find_element_by_id(section_id + ref[i])
        element.send_keys(str(int(section_value[i])))

def valeur(cell_ref):
    return sheet.range(cell_ref).value

# Plage 1
filldata("jYMaRS9Bt2Y",valeur("D12:G12"))

# Plage 2
element = driver.find_element_by_id("dPSO1DDEUj2-HllvX50cXC0-val") 
element.send_keys(str(int(valeur("I12"))))

filldata("ZCiHAAtLEIW",valeur("J12:M12"))
filldata("Xt03XV2wifq",valeur("O12:R12"))
   
element = driver.find_element_by_id("kjq5lD0qaaF-HllvX50cXC0-val")  
element.send_keys(str(int(valeur("T12"))))

element = driver.find_element_by_id("evT9evOWDHM-HllvX50cXC0-val")
element.send_keys(str(int(valeur("U12"))))


# SUBTOTAL OF TX 1
filldata("Raoxru5CpfG",valeur("AK12:AN12"))


# SUBTOTAL OF TX 2
filldata("YSX1bFP33cZ",valeur("AP12:AS12"))


# Fermeture du fichier Excel en cours 
#app.kill()

# IMPORTATION DU FICHIER DE SUIVI

wb = xw.Book("wetmot2.xlsb")

# OUVERTURE DU CLASSEUR ET SELECTION DE FEUILLE CONTENANT LE JEU DE DONNEES
sheet = wb.sheets["RAPPORT_WETMOT_DU_SITE"]


# FERMER LE FIHIER EN ARRIERE PLAN   
#app.kill() 


#element.send_keys(Keys.RETURN) is used to press enter after the values are inserted
#element.send_keys(Keys.RETURN)
#element.close()
