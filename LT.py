import KOMBO_funkce_kusovnik
import os
import time
import openpyxl as excel


file_to_process = "report.txt"
#file_to_process = "demo 2.txt"

path_to_export = 'C:\\Users\\Ondrej.rott\\Documents\\Python\\Lead timy\\Exporty z LN\\' + str(file_to_process)

data_import = KOMBO_funkce_kusovnik.nacteni_dat(path_to_export) # Naceteni dat z txt soboru a rozdeleni na jednotlive linky kusovniku. Vysledek ulozen jako list.
#print(data_import)
kusovnik = KOMBO_funkce_kusovnik.vytvoreni_kusovniku(data_import) # Vezme itemy v kazde lince a udela z nich kusovnik po linkach a ulozi do listu kusovnik  vseho.
#print(kusovnik)
parametry = KOMBO_funkce_kusovnik.databaze_parametru(data_import) # Pro kazdou linku vezme itemy z linky a ulozi jejich hodnoty ex date, lt a ss jako dict do listu all parameters.
#print(parametry)
for kvp in parametry.items():
    if len(kvp[1]) != 11:
        print(kvp)
        print(len(kvp[1]))
data_import.clear()

pocet_linek_kusovniku = len(kusovnik)
vrchol = kusovnik[0][0]
nejdelsi_linka = []
max_lt_itemu = 0
i=0
vysledek = []
multivysledek = []
vsechny_chybejici_routingy = []
vsechny_neplatne_routingy = []
vsechny_vyrabene_use_polozky = []
vsechny_nakupovane_bez_puchase_dat = []
vsechny_nakupovane_bez_puchase_lt = []
vsechny_nakupovane_use_polozky = []
vsechny_itemy_typ_error = []
vsechny_manufactured_dily_bez_kusovniku = []
vrchol_k_zaplanovani = []
vse_k_zaplanovani = []
vrchol_chyby = []

mode_set = False
calculation_mode = ""

while mode_set == False:
    # Zadani jestli chci hromadne doplnovat chybejici routingy a kolik doplnit.
    missing_routing_settings = input(f'Chces hromadne doplnit chybejici routingy jednotnym cilsem? [A]no / [N]e:\n')
    while (missing_routing_settings.upper() != "A") and (missing_routing_settings.upper() != "N"):
        missing_routing_settings = input(f'Zadej bud A pro automaticke doplneni, nebo N pro rucni doplnovani behem programu.\n')
    if missing_routing_settings.upper() == "A":
        while True:
            try:
                chybejici_routingy_equals = int(input(f'Kolik prac. dni budou automaticky doplnene routingy?:\n'))
                if chybejici_routingy_equals > 0:
                    print(f'Bude pocitano {chybejici_routingy_equals} pracovni dni pro kayzdy chybejici routing.\n')
                    break
                else:
                    print(f'Je nutno zadat kladne cislo.')
            except ValueError:
                print(f'Toto neni cislo. Je nutno zadat kladne cislo.')           
    elif missing_routing_settings.upper() == "N":
        chybejici_routingy_equals = 0

    # Zadani jestli chci LT pocitat jako pro nacenovani, nebo jako na MS.
    lt_calculation_mode = input(f'Chces LT pocitat jako pri nacenovani, nebo jako vyroba na MS? [N]acenovani / [MS]MS:\n')
    while (lt_calculation_mode.upper() != "N") and (lt_calculation_mode.upper() != "MS"):
        lt_calculation_mode = input(f'Zadej bud N pro pocitani jako pri nacenovani, nebo MS pro pocitani jako na MS pri vyrobe.\n')
    if lt_calculation_mode.upper() == "N":
        print(f'LT bude pocitano jako pri nacenovani.')
        calculation_mode = "nacenovani"
        mode_set = True          
    elif lt_calculation_mode.upper() == "MS":
        print(f'LT bude pocitano jako na MS pri planovani vyroby.')
        calculation_mode = "ms"
        mode_set = True
    print(calculation_mode)
    print(mode_set)

# Vytvoreni souboru pro zaplanovani itemu.
with open("itemy k zaplanovani.txt", "w") as output_file:

# samotny program
    for line in kusovnik:
        # print(f' RESENA LINKA: {line}')
        if pocet_linek_kusovniku != i+1: #Nejedna se o posledni linku.
            if line[0] == vrchol: #Pokracovani predchoziho vrcholu.

                ###
                lt_linky_output = KOMBO_funkce_kusovnik.lt_linky(line, parametry, chybejici_routingy_equals, calculation_mode) # Vrati [ int(lt linky), list[itemy s chybejicimi routingy]] ]
                KOMBO_funkce_kusovnik.chyby_linky(line, vrchol_chyby, lt_linky_output, vsechny_chybejici_routingy, vsechny_neplatne_routingy, vsechny_itemy_typ_error, vsechny_nakupovane_bez_puchase_dat, vsechny_nakupovane_bez_puchase_lt, vsechny_vyrabene_use_polozky, vsechny_nakupovane_use_polozky, vsechny_manufactured_dily_bez_kusovniku)            

                if lt_linky_output[0] > max_lt_itemu:
                    max_lt_itemu = lt_linky_output[0]
                    nejdelsi_linka = KOMBO_funkce_kusovnik.nejdelsi_linka(line, parametry, calculation_mode)
                
                # Kusovnik k zaplanovani            
                linka_k_zaplanovani = KOMBO_funkce_kusovnik.linka_k_zaplanovani(line, parametry)
                if linka_k_zaplanovani != None:
                    if linka_k_zaplanovani not in vrchol_k_zaplanovani and len(linka_k_zaplanovani) != 0:
                        vrchol_k_zaplanovani.append(linka_k_zaplanovani)
                    if linka_k_zaplanovani not in vse_k_zaplanovani and len(linka_k_zaplanovani) != 0:
                        vse_k_zaplanovani.append(linka_k_zaplanovani)
                # Kusovnik k zaplanovani  
                i+=1
                ###
            else: # Novy vrchol.
                vysledek.append(KOMBO_funkce_kusovnik.vysledek_itemu(nejdelsi_linka, parametry, vrchol, max_lt_itemu, chybejici_routingy_equals, vrchol_chyby, calculation_mode))# Sestaveni nejdelsi linky soucasneho vrcholu a jejiho LT.                  
                if len(vrchol_chyby) != 0:
                    print(f'Vrchol {vrchol} chyby {vrchol_chyby}')     

                for linka in vrchol_k_zaplanovani:
                    output_file.write(str(linka))
                    output_file.write('\n')

                vrchol = line[0] #Resetovani pro novy vrchol.
                max_lt_itemu = 0 #Resetovani pro novy vrchol.
                vrchol_chyby.clear() # Resetovani pro novy vrchol.
                vrchol_k_zaplanovani.clear() # Resetovani pro novy vrchol
                ###
                lt_linky_output = KOMBO_funkce_kusovnik.lt_linky(line, parametry, chybejici_routingy_equals, calculation_mode) # Vrati [ int(lt linky), list[itemy s chybejicimi routingy]] ]
                KOMBO_funkce_kusovnik.chyby_linky(line, vrchol_chyby, lt_linky_output, vsechny_chybejici_routingy, vsechny_neplatne_routingy, vsechny_itemy_typ_error, vsechny_nakupovane_bez_puchase_dat, vsechny_nakupovane_bez_puchase_lt, vsechny_vyrabene_use_polozky, vsechny_nakupovane_use_polozky, vsechny_manufactured_dily_bez_kusovniku)            

                if lt_linky_output[0] > max_lt_itemu:
                    max_lt_itemu = lt_linky_output[0]
                    nejdelsi_linka = KOMBO_funkce_kusovnik.nejdelsi_linka(line, parametry, calculation_mode)

                # Kusovnik k zaplanovani            
                linka_k_zaplanovani = KOMBO_funkce_kusovnik.linka_k_zaplanovani(line, parametry)
                if linka_k_zaplanovani != None:
                    if linka_k_zaplanovani not in vrchol_k_zaplanovani and len(linka_k_zaplanovani) != 0:
                        vrchol_k_zaplanovani.append(linka_k_zaplanovani)
                    if linka_k_zaplanovani not in vse_k_zaplanovani and len(linka_k_zaplanovani) != 0:
                        vse_k_zaplanovani.append(linka_k_zaplanovani)
                # Kusovnik k zaplanovani    

                i+=1
                ###
        else: # Jedna se o posledni linku.
            if line[0] == vrchol: #Pokracovani predchoziho vrcholu. 
                ###
                lt_linky_output = KOMBO_funkce_kusovnik.lt_linky(line, parametry, chybejici_routingy_equals, calculation_mode) # Vrati [ int(lt linky), list[itemy s chybejicimi routingy]] ]
                KOMBO_funkce_kusovnik.chyby_linky(line, vrchol_chyby, lt_linky_output, vsechny_chybejici_routingy, vsechny_neplatne_routingy, vsechny_itemy_typ_error, vsechny_nakupovane_bez_puchase_dat, vsechny_nakupovane_bez_puchase_lt, vsechny_vyrabene_use_polozky, vsechny_nakupovane_use_polozky, vsechny_manufactured_dily_bez_kusovniku)         

                if lt_linky_output[0] > max_lt_itemu:
                    max_lt_itemu = lt_linky_output[0]
                    nejdelsi_linka = KOMBO_funkce_kusovnik.nejdelsi_linka(line, parametry, calculation_mode)

                # Kusovnik k zaplanovani            
                linka_k_zaplanovani = KOMBO_funkce_kusovnik.linka_k_zaplanovani(line, parametry)
                if linka_k_zaplanovani != None:
                    if linka_k_zaplanovani not in vrchol_k_zaplanovani and len(linka_k_zaplanovani) != 0:
                        vrchol_k_zaplanovani.append(linka_k_zaplanovani)
                    if linka_k_zaplanovani not in vse_k_zaplanovani and len(linka_k_zaplanovani) != 0:
                        vse_k_zaplanovani.append(linka_k_zaplanovani)
                # Kusovnik k zaplanovani   
                i+=1
                ###          
                vysledek.append(KOMBO_funkce_kusovnik.vysledek_itemu(nejdelsi_linka, parametry, vrchol, max_lt_itemu, chybejici_routingy_equals, vrchol_chyby, calculation_mode))# Sestaveni nejdelsi linky soucasneho vrcholu a jejiho LT.
                if len(vrchol_chyby) != 0:
                    print(f'Vrchol {vrchol} chyby {vrchol_chyby}')

                for linka in vrchol_k_zaplanovani:
                    output_file.write(str(linka))
                    output_file.write('\n')

                KOMBO_funkce_kusovnik.zaplanovani_do_excelu(vse_k_zaplanovani)

            else: # Novy vrchol.
                vysledek.append(KOMBO_funkce_kusovnik.vysledek_itemu(nejdelsi_linka, parametry, vrchol, max_lt_itemu, chybejici_routingy_equals, vrchol_chyby, calculation_mode))# Sestaveni nejdelsi linky soucasneho vrcholu a jejiho LT.            
                if len(vrchol_chyby) != 0:
                    print(f'Vrchol {vrchol} chyby {vrchol_chyby}')

                for linka in vrchol_k_zaplanovani:
                    output_file.write(str(linka))
                    output_file.write('\n')

                vrchol = line[0] #Resetovani pro novy vrchol.
                max_lt_itemu = 0 #Resetovani pro novy vrchol.
                vrchol_chyby.clear() # Resetovani pro novy vrchol.
                vrchol_k_zaplanovani.clear() # Resetovani pro novy vrchol
                ###
                lt_linky_output = KOMBO_funkce_kusovnik.lt_linky(line, parametry, chybejici_routingy_equals, calculation_mode) # Vrati [ int(lt linky), list[itemy s chybejicimi routingy]] ]
                KOMBO_funkce_kusovnik.chyby_linky(line, vrchol_chyby, lt_linky_output, vsechny_chybejici_routingy, vsechny_neplatne_routingy, vsechny_itemy_typ_error, vsechny_nakupovane_bez_puchase_dat, vsechny_nakupovane_bez_puchase_lt, vsechny_vyrabene_use_polozky, vsechny_nakupovane_use_polozky, vsechny_manufactured_dily_bez_kusovniku)            

                if lt_linky_output[0] > max_lt_itemu:
                    max_lt_itemu = lt_linky_output[0]
                    nejdelsi_linka = KOMBO_funkce_kusovnik.nejdelsi_linka(line, parametry, calculation_mode)

                # Kusovnik k zaplanovani            
                linka_k_zaplanovani = KOMBO_funkce_kusovnik.linka_k_zaplanovani(line, parametry)
                if linka_k_zaplanovani != None:
                    if linka_k_zaplanovani not in vrchol_k_zaplanovani and len(linka_k_zaplanovani) != 0:
                        vrchol_k_zaplanovani.append(linka_k_zaplanovani)
                    if linka_k_zaplanovani not in vse_k_zaplanovani and len(linka_k_zaplanovani) != 0:
                        vse_k_zaplanovani.append(linka_k_zaplanovani)
                # Kusovnik k zaplanovani     
                i+=1
                ###
                vysledek.append(KOMBO_funkce_kusovnik.vysledek_itemu(nejdelsi_linka, parametry, vrchol, max_lt_itemu, chybejici_routingy_equals, vrchol_chyby, calculation_mode))# Sestaveni nejdelsi linky soucasneho vrcholu a jejiho LT.    
                if len(vrchol_chyby) != 0:
                    print(f'Vrchol {vrchol} chyby {vrchol_chyby}')

                for linka in vrchol_k_zaplanovani:
                    output_file.write(str(linka))
                    output_file.write('\n')
                
                KOMBO_funkce_kusovnik.zaplanovani_do_excelu(vse_k_zaplanovani)    

output_file.close()

print("\n")

print("SALES LT (kalendarni dny):\n")

for item in vysledek:
    print(item)
print("\n")


if len(vsechny_chybejici_routingy) != 0:
    print("Seznam vsech chybejicich routingu z reportu:")

    for item in vsechny_chybejici_routingy:
        if item[0:3] == "PMP":
            print(f'{item[0:9]}:{item[9:len(item)]}:{parametry.get(item).get("description")}')
        elif parametry.get(vrchol).get("supplier") == "I00000008":    
            print(f'Lamphun_A:{item}:{parametry.get(item).get("description")}')
        else:
            print(f'anonymni_:{item}:{parametry.get(item).get("description")}')


if len(vsechny_neplatne_routingy) != 0:
    print("\nSeznam vsech dilu s neplatnym routingem z reportu:")
    for item in vsechny_neplatne_routingy:
        if item[0:3] == "PMP":
            print(item[0:9]+":"+str(item[9:len(item)]))
        elif parametry.get(vrchol).get("supplier") == "I00000008":    
            print("Lamphun_A:"+str(item))
        else:
            print("anonymni_:"+str(item))
         

if len(vsechny_nakupovane_bez_puchase_dat) != 0:
    print("\nSeznam vsech nakupovanych dilu bez dodavatele nebo nakupciho z reportu:")
    for item in vsechny_nakupovane_bez_puchase_dat:
            print("anonymni_:"+str(item))


if len(vsechny_nakupovane_bez_puchase_lt) != 0:
    print("\nSeznam vsech nakupovanych dilu bez purchased LT v LN z reportu:")
    for item in vsechny_nakupovane_bez_puchase_lt:
            print("anonymni_:"+str(item))

if len(vsechny_itemy_typ_error) != 0:
    print("\nSeznam vsech dilu s neznamym typem z reportu:")
    for item in vsechny_itemy_typ_error:
            print("anonymni_:"+str(item))

if len(vsechny_vyrabene_use_polozky) != 0:
    print("\nSeznam vsech vyrabenych USE polozek z reportu:")
    for item in vsechny_vyrabene_use_polozky:
        if item[0:3] == "PMP":
            print(item[0:9]+":"+str(item[9:len(item)]))
        elif parametry.get(vrchol).get("supplier") == "I00000008":    
            print("Lamphun_A:"+str(item))
        else:
            print("anonymni_:"+str(item))

if len(vsechny_nakupovane_use_polozky) != 0:
    print("\nSeznam nakupovanych USE polozek z reportu:")
    for item in vsechny_nakupovane_use_polozky:
            print("anonymni_:"+str(item)) 

if len(vsechny_manufactured_dily_bez_kusovniku) != 0:
    print("\nSeznam Manufactured dilu bez kusovniku:")
    for item in vsechny_manufactured_dily_bez_kusovniku:
        if item[0:3] == "PMP":
            print(item[0:9]+":"+str(item[9:len(item)]))
        elif parametry.get(vrchol).get("supplier") == "I00000008":    
            print("Lamphun_A:"+str(item))
        else:
            print("anonymni_:"+str(item))     

vsechny_manufactured_dily_bez_kusovniku
print("\nKONEC PROGRAMU")