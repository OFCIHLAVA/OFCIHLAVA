import moduly.funkce.funkce_prace as fce
import sys

# Soubor vyexportovanych kusovniku z CQ.
file_update ='Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\SFExBFExMIX\\databaze\\Exporty LN\\kusovniky_update.txt'

print(f'START . . .\n')

### Priprava dat na dalsi zpracovani.

#Import souboru.
try:
    import_update = fce.import_kusovniku_LN(file_update)
    print(f'Soubor kusovníků boud z CQ dat načten ({file_update}) . . .\n')
except FileNotFoundError:      
    print(f'ERROR:\nVe složce Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\SFExBFExMIX\\databaze\\Exporty LN nebyl nalezen soubor pro update "kusoviky_update.txt".\nZkontrolouj umístění souboru a jméno souboru...')
    input(f'\nPress ENTRER to EXIT program . . .')
    sys.exit(1)

#  Vytvoreni seznamu boud z CQ exportu, ktere chceme pridat do stavajici databaze.
boudy_update = fce.seznam_boud_z_importu(import_update)
print(f'Vytvořen seznam boud z CQ dat, které chceme přidat do databáze . . .\n')
print(f'Počet boud, které chceme přidat: {len(boudy_update)}.')

# Vytvoreni slovniku boud z CQ exportu s jejich kusovniky.
obsah_databaze_kusovniku_update = fce.boudy_s_kusovnikem(boudy_update, import_update)
dict_databaze_update = obsah_databaze_kusovniku_update[0]
print(f'Vytvořena databáze boud, které chceme přidat do databáze z CQ dat s jejich kusovníky . . .\n')

### Nacteni stavajici databaze.
current_database = fce.nacteni_databaze_boud_pro_dotaz("Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\SFExBFExMIX\\databaze\\databaze boud s kusovniky.txt")
# print(current_database)
print(f'Načtení stávající databáze pro porovnání . . .')
print(f'Počet boud ve stávající databázi PŘED updatem: {len(current_database)}.\n')
# print(current_database)

### Kontrola, zda jsou boudy, ktere chceme updatovat ve stavajici databazi. Pokud jsou, ale neshoduji se → Prepisou se. Pokud nejsou → Pridaji se. Pokud uz jsou tam stejna data, nic se nestane.

print(f'Porovnávání stávající databáze s boudami, které chceme přidat . . .\n')

for nova_bouda in boudy_update:
    if nova_bouda in current_database:
        if dict_databaze_update.get(nova_bouda) == current_database.get(nova_bouda):
            print(f'Nová bouda: {nova_bouda}', f'JE ve stávající databázi a kusovníky JSOU totožné →→→ OK, není třeba updatovat.')
        else:
            print(f'Nová bouda: {nova_bouda}', f'JE ve stávající databázi ale kusovníky NEJSOU totožné →→→ NOK, UPDATOVAT.')
            current_database[nova_bouda] = dict_databaze_update.get(nova_bouda)
    else:
        print(f'Nová bouda: {nova_bouda}', f'NENÍ ve stávající databázi →→→ NOK, UPDATOVAT.')
        current_database[nova_bouda] = dict_databaze_update.get(nova_bouda)
print(f'Počet boud ve nové databázi PO updatu: {len(current_database)}.\n')

### Priprava slovniku updatovane databaze na ulozeni do txt souboru databaze boud s kusovniky.

seznam_vseho_str = str(current_database)
na_linky = seznam_vseho_str.replace("], ", "<+++>")
rozdeleni_pn = na_linky.replace("', '", "|")
apostrof = rozdeleni_pn.replace("'", "")
leva_hranata = apostrof.replace("[", "")
prava_hranata = leva_hranata.replace("]", "")
mezera = prava_hranata.replace(" ", "")
leva_slozena = mezera.replace("{", "")
prava_slozena = leva_slozena.replace("}", "")

print(f'Databáze po updatu připravena na uložení do txt soubotu . . .\n')

# Vysledek.
str_dict_seznam_vseho = prava_slozena

# Vysledek se ulozi do souboru POST UPDATE databaze. Pokud ji chceme pouzivat, je touto treba nahradit stavajici databazi.

print(f'Zapisování Databáze po updatu do txt souboru "POST_UPDATE_programy_databaze_update.txt" ve složce Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\SFExBFExMIX\\databaze . . .\n')
with open("Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\SFExBFExMIX\\databaze\\POST_UPDATE_programy_databaze_update.txt", "w") as new_updated_database:
    new_updated_database.write(str_dict_seznam_vseho)
    new_updated_database.close()

print(f'Databáze úspěšně updatována . . .\n')

input(f'Press ENTRER to EXIT program . . .')