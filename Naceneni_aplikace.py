from Programy import funkce_prace
from Leadtime import KOMBO_funkce_kusovnik
import os
import time
import openpyxl as excel
from operator import le


file_to_process = "report.txt"
path_to_export = 'C:\\Users\\Ondrej.rott\\Documents\\Python\\Nove nacenovani\\LT LN export\\' + str(file_to_process)
krok = pocet_sloupcu_dat_v_exportu_kusovniku_ze_cq = 24

# CQ import nacteni dat kusovniku priprava
data_import = KOMBO_funkce_kusovnik.nacteni_dat(path_to_export) # Naceteni dat z txt soboru a rozdeleni na jednotlive linky kusovniku. Vysledek ulozen jako list.
# print(data_import)
parametry = KOMBO_funkce_kusovnik.databaze_parametru(data_import, krok) # Pro kazdou linku vezme itemy z linky a ulozi jejich hodnoty ex date, lt a ss jako dict do listu all parameters.
# print(parametry)
for kvp in parametry.items(): # Kontrola, ze vsechny polozky z reportu maji vyplnene v parametrech vsechny hodnoty.
    if len(kvp[1]) != (krok-1):
        print(kvp)
        print(len(kvp[1]))
kusovnik, kusovnik_bom_qty = KOMBO_funkce_kusovnik.vytvoreni_kusovniku(data_import, krok, parametry) # Vezme itemy v kazde lince a udela z nich kusovnik po linkach a ulozi do listu kusovnik  vseho.
# for line in kusovnik:
#     print(line)
# for line in kusovnik_bom_qty:
#     print(line)
data_import.clear()

# Priprava databaze pro zjistovani programu SFE/BFE/MIX
path_kusovniky_databaze = 'C:\\Users\\Ondrej.rott\\Documents\\Python\\Nove nacenovani\\Programy\\databaze boud s kusovniky.txt'
path_program_databaze = 'C:\\Users\\Ondrej.rott\\Documents\\Python\\Nove nacenovani\\Programy\\seznam programu.txt'

databaze_pro_dotaz_programy = funkce_prace.nacteni_databaze_boud_pro_dotaz(path_kusovniky_databaze)
kvp_programy = funkce_prace.programy_boud(path_program_databaze)

# Dalsi veci priprava
pocet_linek_kusovniku = len(kusovnik)
vrchol = kusovnik[0][0]
nejdelsi_linka = []
max_lt_itemu = 0
i=0
vysledek = []
multivysledek = []

# Chyby
vsechny_chybejici_routingy = []
vsechny_neplatne_routingy = []
vsechny_vyrabene_use_polozky = []
vsechny_nakupovane_bez_puchase_dat = []
vsechny_nakupovane_bez_puchase_lt = []
vsechny_nakupovane_use_polozky = []
vsechny_itemy_typ_error = []
vsechny_manufactured_dily_bez_kusovniku = []
vsechny_vyrabene_bom_qty_0 = []

vrchol_chyby = []

# Zaplanovani do txt. souboru
vrchol_k_zaplanovani = []
vse_k_zaplanovani = []
vsechny_manufactured_polozky_pod_projekt = []

# Naceneni do tabulek priprava
uz_nacenene_vrcholy = []
sfe_tabulka = []
bfe_tabulka = []
neznamy_program_tabulka = []
zatim_nenacenitelne_itemy_tabulka = []

# Zadavani platnych kalkulacnich projektu do nastaveni
with open('C:\\Users\\Ondrej.rott\\Documents\\Python\\Nove nacenovani\\nastaveni\\config.txt', 'r') as config:
    nastaveni_data = config.readlines()
    config.close()
config_dict = {}
for line in nastaveni_data:
    config_dict[line.split(" = ")[0]] = line.split(" = ")[1].replace("\n","").replace("'", "").replace("[","").replace("]","").replace(" ","")
platne_kalkulacni_projekty = config_dict.get("aktualne platne kalkulacni projekty").split(",")
# print(platne_kalkulacni_projekty)


# Samotny program
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
        print(f'LT bude pocitano jako pri nacenovani.\n')
        calculation_mode = "nacenovani"
        mode_set = True          
    elif lt_calculation_mode.upper() == "MS":
        print(f'LT bude pocitano jako na MS pri planovani vyroby.\n')
        calculation_mode = "ms"

    # Overeni a pripadne aktualizovani seznamu nacenovacich projektu.
    print(f'Projekty {platne_kalkulacni_projekty} jsou nastaveny jako platne kalkulacni projekty.')
    kalkulacni_projekty_zmenit = input(f'Je tento seznam v poradku? [A]no. / Seznam uz neni aktualni. Zadat jiny seznam [N]e:\n')

    while (kalkulacni_projekty_zmenit.upper() != "A") and (kalkulacni_projekty_zmenit.upper() != "N"):
        kalkulacni_projekty_zmenit = input(f'Zadej bud A pro pokracovani se stavajicimi kalk. projekty {platne_kalkulacni_projekty}, nebo N pro zadani noveho seznamu.\n')
    if kalkulacni_projekty_zmenit.upper() == "A":
        print(f'Seznam kalkulacnich projektu se nemeni {platne_kalkulacni_projekty}.')
        mode_set = True          
    elif kalkulacni_projekty_zmenit.upper() == "N":
        ok = False
        while not ok:
            nove_zadane_kalkulacni_projekty = input(f'Zadej novy seznam platnych kalkulacnich projektu. Pri vice projektech je oddel carkou (,): \n') 
            nove_zadane_kalkulacni_projekty = [(projekt).strip() for projekt in nove_zadane_kalkulacni_projekty.split(",")]                   
            # overeni platne zadanycho projektu 
            check = []
            for projekt in nove_zadane_kalkulacni_projekty:
                if len(projekt) == 9:
                    delka = 'OK'
                    check.append(delka)
                else:
                    delka = 'error'
                    check.append(delka)               
                if projekt[0:3].upper() == "PMP":
                    nove_zadane_kalkulacni_projekty[nove_zadane_kalkulacni_projekty.index(projekt)] = f'{projekt[0:3].upper()}{projekt[3:9]}'
                    prefix = 'OK'
                    check.append(prefix)
                else:
                    prefix = 'error'
                    check.append(prefix)   
                for znak in projekt[3:9]:
                    try:
                        int(znak)
                        cislo = 'OK'
                        check.append(cislo)
                    except:    
                        cislo = 'error'
                        check.append(cislo)
            if "error" not in check:
                ok = True
            else:
                ok = False
                print(f'Projekty musi mit 9 znaku a obsahovat PMP + 6cisel.')
        with open('C:\\Users\\Ondrej.rott\\Documents\\Python\\Nove nacenovani\\nastaveni\\config.txt', 'w') as config:
            s = str(nove_zadane_kalkulacni_projekty)
            config.write(f'aktualne platne kalkulacni projekty = {s}')
            config.close()
        print(f'Nove platne kalkulacni projekty: {nove_zadane_kalkulacni_projekty}.')
        platne_kalkulacni_projekty = nove_zadane_kalkulacni_projekty
        mode_set = True
    print(f'Kalkulacni mod: {calculation_mode}.')
    # print(mode_set)

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
                linka_k_zaplanovani = KOMBO_funkce_kusovnik.linka_k_zaplanovani(line, parametry, vsechny_manufactured_polozky_pod_projekt, platne_kalkulacni_projekty)
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
                
                ### Naceneni vysledku do nacenovacich tabulek
                if calculation_mode == "nacenovani":
                # 1. Pokud se jedna o anonymni ciste nakupovanou vec ne z lamphunu. → rovnou nacenit do tabulky 
                    if vrchol[0:3] != "PMP" and parametry.get(vrchol).get("supplier") != "0" and parametry.get(vrchol).get("supplier") != "I00000008" and parametry.get(vrchol).get("nakupci") != "0" and parametry.get(vrchol).get("nakupci") != "PZP001":                 
                        # print(f'ciste ano naku polozka {vrchol}')
                        program = funkce_prace.dotaz_pn_program(vrchol, databaze_pro_dotaz_programy, kvp_programy)[0].split(":")[1]                   
                        # print(f'vrchol {vrchol} - {program}')
                        naceneny_item, bom_qty_check = KOMBO_funkce_kusovnik.naceneni_do_tabulek(vrchol, parametry, program, kusovnik, kusovnik_bom_qty)
                        # zachyceni vyrabenych itemu s 0 bom qty / PLACARDU s kusovnikem a pridani do chyb.
                        for item in bom_qty_check[1]:
                            if item not in vsechny_vyrabene_bom_qty_0:
                                vsechny_vyrabene_bom_qty_0.append(item)

                        if program in ("SFE", "MIX"):
                            sfe_tabulka.append(naceneny_item)
                        elif program == "BFE":
                            bfe_tabulka.append(naceneny_item)
                        elif program not in ("SFE", "MIX", "BFE"):
                            neznamy_program_tabulka.append(naceneny_item)
                        uz_nacenene_vrcholy.append(vrchol)
                        if vrchol in zatim_nenacenitelne_itemy_tabulka:
                            zatim_nenacenitelne_itemy_tabulka.remove(vrchol)    
                # 2. Pokud se jedna o item pod platnym kalkulacnim projektem → nacenit do tabulky
                    elif vrchol[0:9] in platne_kalkulacni_projekty:
                        # print(f'vyrabena pod platnym kalkulakem {vrchol}')
                        program = funkce_prace.dotaz_pn_program(vrchol[9:len(vrchol)], databaze_pro_dotaz_programy, kvp_programy)[0].split(":")[1]    
                        # print(f'vrchol {vrchol} - {program}')
                        naceneny_item, bom_qty_check = KOMBO_funkce_kusovnik.naceneni_do_tabulek(vrchol, parametry, program, kusovnik, kusovnik_bom_qty)
                        # zachyceni vyrabenych itemu s 0 bom qty / PLACARDU s kusovnikem a pridani do chyb.
                        for item in bom_qty_check[1]:
                            if item not in vsechny_vyrabene_bom_qty_0:
                                vsechny_vyrabene_bom_qty_0.append(item)

                        if program in ("SFE", "MIX"):
                            sfe_tabulka.append(naceneny_item)
                        elif program == "BFE":
                            bfe_tabulka.append(naceneny_item)
                        elif program not in ("SFE", "MIX", "BFE"):
                            neznamy_program_tabulka.append(naceneny_item)
                        uz_nacenene_vrcholy.append(vrchol[9:len(vrchol)])
                        if vrchol[9:len(vrchol)] in zatim_nenacenitelne_itemy_tabulka:
                            zatim_nenacenitelne_itemy_tabulka.remove(vrchol[9:len(vrchol)])                         
                # 3. Pokud se jedna o MAN PLACARD / ID PLACARD → naceni do tabulky           
                    elif (KOMBO_funkce_kusovnik.je_to_man_placard(vrchol, parametry) or KOMBO_funkce_kusovnik.je_to_id_placard(vrchol, parametry)):
                        if (vrchol[0:3] != "PMP" and vrchol not in uz_nacenene_vrcholy) or (vrchol[0:3] == "PMP" and vrchol[9:len(vrchol)] not in uz_nacenene_vrcholy):                            
                            # print(f'MAN PLACARD / ID PLACARD {vrchol}')
                            # Pokud se jedna o PLACARD pod PMP projektem
                            if vrchol[0:3] == "PMP":
                                program = funkce_prace.dotaz_pn_program(vrchol[9:len(vrchol)], databaze_pro_dotaz_programy, kvp_programy)[0].split(":")[1]
                            else:                           
                                program = funkce_prace.dotaz_pn_program(vrchol, databaze_pro_dotaz_programy, kvp_programy)[0].split(":")[1]
                            # print(f'vrchol {vrchol} - {program}')
                            naceneny_item, bom_qty_check = KOMBO_funkce_kusovnik.naceneni_do_tabulek(vrchol, parametry, program, kusovnik, kusovnik_bom_qty)
                            # zachyceni vyrabenych itemu s 0 bom qty / PLACARDU s kusovnikem a pridani do chyb.
                            for item in bom_qty_check[1]:
                                if item not in vsechny_vyrabene_bom_qty_0:
                                    vsechny_vyrabene_bom_qty_0.append(item)

                            if program in ("SFE", "MIX"):
                                sfe_tabulka.append(naceneny_item)
                            elif program == "BFE":
                                bfe_tabulka.append(naceneny_item)
                            elif program not in ("SFE", "MIX", "BFE"):
                                neznamy_program_tabulka.append(naceneny_item)
                            if vrchol[0:3] != "PMP":
                                uz_nacenene_vrcholy.append(vrchol)
                                if vrchol in zatim_nenacenitelne_itemy_tabulka:
                                    zatim_nenacenitelne_itemy_tabulka.remove(vrchol)                                 
                            elif vrchol[0:3] == "PMP":
                                uz_nacenene_vrcholy.append(vrchol[9:len(vrchol)])
                                if vrchol[9:len(vrchol)] in zatim_nenacenitelne_itemy_tabulka:
                                    zatim_nenacenitelne_itemy_tabulka.remove(vrchol[9:len(vrchol)])                                  
                    else:
                        if vrchol[0:3] == "PMP":
                            if vrchol[9:len(vrchol)] not in uz_nacenene_vrcholy:
                                zatim_nenacenitelne_itemy_tabulka.append(vrchol)
                        else:                        
                            if vrchol not in uz_nacenene_vrcholy:
                                zatim_nenacenitelne_itemy_tabulka.append(vrchol)

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
                linka_k_zaplanovani = KOMBO_funkce_kusovnik.linka_k_zaplanovani(line, parametry, vsechny_manufactured_polozky_pod_projekt, platne_kalkulacni_projekty)
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
                linka_k_zaplanovani = KOMBO_funkce_kusovnik.linka_k_zaplanovani(line, parametry, vsechny_manufactured_polozky_pod_projekt, platne_kalkulacni_projekty)
                if linka_k_zaplanovani != None:
                    if linka_k_zaplanovani not in vrchol_k_zaplanovani and len(linka_k_zaplanovani) != 0:
                        vrchol_k_zaplanovani.append(linka_k_zaplanovani)
                    if linka_k_zaplanovani not in vse_k_zaplanovani and len(linka_k_zaplanovani) != 0:
                        vse_k_zaplanovani.append(linka_k_zaplanovani)
                # Kusovnik k zaplanovani   
                i+=1
                ###          
                vysledek.append(KOMBO_funkce_kusovnik.vysledek_itemu(nejdelsi_linka, parametry, vrchol, max_lt_itemu, chybejici_routingy_equals, vrchol_chyby, calculation_mode))# Sestaveni nejdelsi linky soucasneho vrcholu a jejiho LT.
                
                ### Naceneni vysledku do nacenovacich tabulek
                if calculation_mode == "nacenovani":
                # 1. Pokud se jedna o anonymni ciste nakupovanou vec ne z lamphunu. → rovnou nacenit do tabulky 
                    if vrchol[0:3] != "PMP" and parametry.get(vrchol).get("supplier") != "0" and parametry.get(vrchol).get("supplier") != "I00000008" and parametry.get(vrchol).get("nakupci") != "0" and parametry.get(vrchol).get("nakupci") != "PZP001":                 
                        # print(f'ciste ano naku polozka {vrchol}')
                        program = funkce_prace.dotaz_pn_program(vrchol, databaze_pro_dotaz_programy, kvp_programy)[0].split(":")[1]                   
                        # print(f'vrchol {vrchol} - {program}')
                        naceneny_item, bom_qty_check = KOMBO_funkce_kusovnik.naceneni_do_tabulek(vrchol, parametry, program, kusovnik, kusovnik_bom_qty)
                        # zachyceni vyrabenych itemu s 0 bom qty / PLACARDU s kusovnikem a pridani do chyb.
                        for item in bom_qty_check[1]:
                            if item not in vsechny_vyrabene_bom_qty_0:
                                vsechny_vyrabene_bom_qty_0.append(item)

                        if program in ("SFE", "MIX"):
                            sfe_tabulka.append(naceneny_item)
                        elif program == "BFE":
                            bfe_tabulka.append(naceneny_item)
                        elif program not in ("SFE", "MIX", "BFE"):
                            neznamy_program_tabulka.append(naceneny_item)
                        uz_nacenene_vrcholy.append(vrchol)
                        if vrchol in zatim_nenacenitelne_itemy_tabulka:
                            zatim_nenacenitelne_itemy_tabulka.remove(vrchol)    
                # 2. Pokud se jedna o item pod platnym kalkulacnim projektem → nacenit do tabulky
                    elif vrchol[0:9] in platne_kalkulacni_projekty:
                        # print(f'vyrabena pod platnym kalkulakem {vrchol}')
                        program = funkce_prace.dotaz_pn_program(vrchol[9:len(vrchol)], databaze_pro_dotaz_programy, kvp_programy)[0].split(":")[1]    
                        # print(f'vrchol {vrchol} - {program}')
                        naceneny_item, bom_qty_check = KOMBO_funkce_kusovnik.naceneni_do_tabulek(vrchol, parametry, program, kusovnik, kusovnik_bom_qty)
                        # zachyceni vyrabenych itemu s 0 bom qty / PLACARDU s kusovnikem a pridani do chyb.
                        for item in bom_qty_check[1]:
                            if item not in vsechny_vyrabene_bom_qty_0:
                                vsechny_vyrabene_bom_qty_0.append(item)

                        if program in ("SFE", "MIX"):
                            sfe_tabulka.append(naceneny_item)
                        elif program == "BFE":
                            bfe_tabulka.append(naceneny_item)
                        elif program not in ("SFE", "MIX", "BFE"):
                            neznamy_program_tabulka.append(naceneny_item)
                        uz_nacenene_vrcholy.append(vrchol[9:len(vrchol)])
                        if vrchol[9:len(vrchol)] in zatim_nenacenitelne_itemy_tabulka:
                            zatim_nenacenitelne_itemy_tabulka.remove(vrchol[9:len(vrchol)])                         
                # 3. Pokud se jedna o MAN PLACARD / ID PLACARD → naceni do tabulky           
                    elif (KOMBO_funkce_kusovnik.je_to_man_placard(vrchol, parametry) or KOMBO_funkce_kusovnik.je_to_id_placard(vrchol, parametry)):
                        if (vrchol[0:3] != "PMP" and vrchol not in uz_nacenene_vrcholy) or (vrchol[0:3] == "PMP" and vrchol[9:len(vrchol)] not in uz_nacenene_vrcholy):                            
                            # print(f'MAN PLACARD / ID PLACARD {vrchol}')
                            # Pokud se jedna o PLACARD pod PMP projektem
                            if vrchol[0:3] == "PMP":
                                program = funkce_prace.dotaz_pn_program(vrchol[9:len(vrchol)], databaze_pro_dotaz_programy, kvp_programy)[0].split(":")[1]
                            else:                           
                                program = funkce_prace.dotaz_pn_program(vrchol, databaze_pro_dotaz_programy, kvp_programy)[0].split(":")[1]
                            # print(f'vrchol {vrchol} - {program}')
                            naceneny_item, bom_qty_check = KOMBO_funkce_kusovnik.naceneni_do_tabulek(vrchol, parametry, program, kusovnik, kusovnik_bom_qty)
                            # zachyceni vyrabenych itemu s 0 bom qty / PLACARDU s kusovnikem a pridani do chyb.
                            for item in bom_qty_check[1]:
                                if item not in vsechny_vyrabene_bom_qty_0:
                                    vsechny_vyrabene_bom_qty_0.append(item)

                            if program in ("SFE", "MIX"):
                                sfe_tabulka.append(naceneny_item)
                            elif program == "BFE":
                                bfe_tabulka.append(naceneny_item)
                            elif program not in ("SFE", "MIX", "BFE"):
                                neznamy_program_tabulka.append(naceneny_item)
                            if vrchol[0:3] != "PMP":
                                uz_nacenene_vrcholy.append(vrchol)
                                if vrchol in zatim_nenacenitelne_itemy_tabulka:
                                    zatim_nenacenitelne_itemy_tabulka.remove(vrchol)                                 
                            elif vrchol[0:3] == "PMP":
                                uz_nacenene_vrcholy.append(vrchol[9:len(vrchol)])
                                if vrchol[9:len(vrchol)] in zatim_nenacenitelne_itemy_tabulka:
                                    zatim_nenacenitelne_itemy_tabulka.remove(vrchol[9:len(vrchol)])                                  
                    else:
                        if vrchol[0:3] == "PMP":
                            if vrchol[9:len(vrchol)] not in uz_nacenene_vrcholy:
                                zatim_nenacenitelne_itemy_tabulka.append(vrchol)
                        else:                        
                            if vrchol not in uz_nacenene_vrcholy:
                                zatim_nenacenitelne_itemy_tabulka.append(vrchol)

                if len(vrchol_chyby) != 0:
                    print(f'Vrchol {vrchol} chyby {vrchol_chyby}')

                for linka in vrchol_k_zaplanovani:
                    output_file.write(str(linka))
                    output_file.write('\n')

                # KOMBO_funkce_kusovnik.zaplanovani_do_excelu(vse_k_zaplanovani)

            else: # Novy vrchol.
                vysledek.append(KOMBO_funkce_kusovnik.vysledek_itemu(nejdelsi_linka, parametry, vrchol, max_lt_itemu, chybejici_routingy_equals, vrchol_chyby, calculation_mode))# Sestaveni nejdelsi linky soucasneho vrcholu a jejiho LT.            
                
                ### Naceneni vysledku do nacenovacich tabulek
                if calculation_mode == "nacenovani":
                # 1. Pokud se jedna o anonymni ciste nakupovanou vec ne z lamphunu. → rovnou nacenit do tabulky 
                    if vrchol[0:3] != "PMP" and parametry.get(vrchol).get("supplier") != "0" and parametry.get(vrchol).get("supplier") != "I00000008" and parametry.get(vrchol).get("nakupci") != "0" and parametry.get(vrchol).get("nakupci") != "PZP001":                 
                        # print(f'ciste ano naku polozka {vrchol}')
                        program = funkce_prace.dotaz_pn_program(vrchol, databaze_pro_dotaz_programy, kvp_programy)[0].split(":")[1]                   
                        # print(f'vrchol {vrchol} - {program}')
                        naceneny_item, bom_qty_check = KOMBO_funkce_kusovnik.naceneni_do_tabulek(vrchol, parametry, program, kusovnik, kusovnik_bom_qty)
                        # zachyceni vyrabenych itemu s 0 bom qty / PLACARDU s kusovnikem a pridani do chyb.
                        for item in bom_qty_check[1]:
                            if item not in vsechny_vyrabene_bom_qty_0:
                                vsechny_vyrabene_bom_qty_0.append(item)

                        if program in ("SFE", "MIX"):
                            sfe_tabulka.append(naceneny_item)
                        elif program == "BFE":
                            bfe_tabulka.append(naceneny_item)
                        elif program not in ("SFE", "MIX", "BFE"):
                            neznamy_program_tabulka.append(naceneny_item)
                        uz_nacenene_vrcholy.append(vrchol)
                        if vrchol in zatim_nenacenitelne_itemy_tabulka:
                            zatim_nenacenitelne_itemy_tabulka.remove(vrchol)    
                # 2. Pokud se jedna o item pod platnym kalkulacnim projektem → nacenit do tabulky
                    elif vrchol[0:9] in platne_kalkulacni_projekty:
                        # print(f'vyrabena pod platnym kalkulakem {vrchol}')
                        program = funkce_prace.dotaz_pn_program(vrchol[9:len(vrchol)], databaze_pro_dotaz_programy, kvp_programy)[0].split(":")[1]    
                        # print(f'vrchol {vrchol} - {program}')
                        naceneny_item, bom_qty_check = KOMBO_funkce_kusovnik.naceneni_do_tabulek(vrchol, parametry, program, kusovnik, kusovnik_bom_qty)
                        # zachyceni vyrabenych itemu s 0 bom qty / PLACARDU s kusovnikem a pridani do chyb.
                        for item in bom_qty_check[1]:
                            if item not in vsechny_vyrabene_bom_qty_0:
                                vsechny_vyrabene_bom_qty_0.append(item)

                        if program in ("SFE", "MIX"):
                            sfe_tabulka.append(naceneny_item)
                        elif program == "BFE":
                            bfe_tabulka.append(naceneny_item)
                        elif program not in ("SFE", "MIX", "BFE"):
                            neznamy_program_tabulka.append(naceneny_item)
                        uz_nacenene_vrcholy.append(vrchol[9:len(vrchol)])
                        if vrchol[9:len(vrchol)] in zatim_nenacenitelne_itemy_tabulka:
                            zatim_nenacenitelne_itemy_tabulka.remove(vrchol[9:len(vrchol)])                         
                # 3. Pokud se jedna o MAN PLACARD / ID PLACARD → naceni do tabulky           
                    elif (KOMBO_funkce_kusovnik.je_to_man_placard(vrchol, parametry) or KOMBO_funkce_kusovnik.je_to_id_placard(vrchol, parametry)):
                        if (vrchol[0:3] != "PMP" and vrchol not in uz_nacenene_vrcholy) or (vrchol[0:3] == "PMP" and vrchol[9:len(vrchol)] not in uz_nacenene_vrcholy):                            
                            # print(f'MAN PLACARD / ID PLACARD {vrchol}')
                            # Pokud se jedna o PLACARD pod PMP projektem
                            if vrchol[0:3] == "PMP":
                                program = funkce_prace.dotaz_pn_program(vrchol[9:len(vrchol)], databaze_pro_dotaz_programy, kvp_programy)[0].split(":")[1]
                            else:                           
                                program = funkce_prace.dotaz_pn_program(vrchol, databaze_pro_dotaz_programy, kvp_programy)[0].split(":")[1]
                            # print(f'vrchol {vrchol} - {program}')
                            naceneny_item, bom_qty_check = KOMBO_funkce_kusovnik.naceneni_do_tabulek(vrchol, parametry, program, kusovnik, kusovnik_bom_qty)
                            # zachyceni vyrabenych itemu s 0 bom qty / PLACARDU s kusovnikem a pridani do chyb.
                            for item in bom_qty_check[1]:
                                if item not in vsechny_vyrabene_bom_qty_0:
                                    vsechny_vyrabene_bom_qty_0.append(item)

                            if program in ("SFE", "MIX"):
                                sfe_tabulka.append(naceneny_item)
                            elif program == "BFE":
                                bfe_tabulka.append(naceneny_item)
                            elif program not in ("SFE", "MIX", "BFE"):
                                neznamy_program_tabulka.append(naceneny_item)
                            if vrchol[0:3] != "PMP":
                                uz_nacenene_vrcholy.append(vrchol)
                                if vrchol in zatim_nenacenitelne_itemy_tabulka:
                                    zatim_nenacenitelne_itemy_tabulka.remove(vrchol)                                 
                            elif vrchol[0:3] == "PMP":
                                uz_nacenene_vrcholy.append(vrchol[9:len(vrchol)])
                                if vrchol[9:len(vrchol)] in zatim_nenacenitelne_itemy_tabulka:
                                    zatim_nenacenitelne_itemy_tabulka.remove(vrchol[9:len(vrchol)])                                  
                    else:
                        if vrchol[0:3] == "PMP":
                            if vrchol[9:len(vrchol)] not in uz_nacenene_vrcholy:
                                zatim_nenacenitelne_itemy_tabulka.append(vrchol)
                        else:                        
                            if vrchol not in uz_nacenene_vrcholy:
                                zatim_nenacenitelne_itemy_tabulka.append(vrchol)

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
                linka_k_zaplanovani = KOMBO_funkce_kusovnik.linka_k_zaplanovani(line, parametry, vsechny_manufactured_polozky_pod_projekt, platne_kalkulacni_projekty)
                if linka_k_zaplanovani != None:
                    if linka_k_zaplanovani not in vrchol_k_zaplanovani and len(linka_k_zaplanovani) != 0:
                        vrchol_k_zaplanovani.append(linka_k_zaplanovani)
                    if linka_k_zaplanovani not in vse_k_zaplanovani and len(linka_k_zaplanovani) != 0:
                        vse_k_zaplanovani.append(linka_k_zaplanovani)
                # Kusovnik k zaplanovani     
                i+=1
                ###
                vysledek.append(KOMBO_funkce_kusovnik.vysledek_itemu(nejdelsi_linka, parametry, vrchol, max_lt_itemu, chybejici_routingy_equals, vrchol_chyby, calculation_mode))# Sestaveni nejdelsi linky soucasneho vrcholu a jejiho LT.    
                
                ### Naceneni vysledku do nacenovacich tabulek
                if calculation_mode == "nacenovani":
                # 1. Pokud se jedna o anonymni ciste nakupovanou vec ne z lamphunu. → rovnou nacenit do tabulky 
                    if vrchol[0:3] != "PMP" and parametry.get(vrchol).get("supplier") != "0" and parametry.get(vrchol).get("supplier") != "I00000008" and parametry.get(vrchol).get("nakupci") != "0" and parametry.get(vrchol).get("nakupci") != "PZP001":                 
                        # print(f'ciste ano naku polozka {vrchol}')
                        program = funkce_prace.dotaz_pn_program(vrchol, databaze_pro_dotaz_programy, kvp_programy)[0].split(":")[1]                   
                        # print(f'vrchol {vrchol} - {program}')
                        naceneny_item, bom_qty_check = KOMBO_funkce_kusovnik.naceneni_do_tabulek(vrchol, parametry, program, kusovnik, kusovnik_bom_qty)
                        # zachyceni vyrabenych itemu s 0 bom qty / PLACARDU s kusovnikem a pridani do chyb.
                        for item in bom_qty_check[1]:
                            if item not in vsechny_vyrabene_bom_qty_0:
                                vsechny_vyrabene_bom_qty_0.append(item)

                        if program in ("SFE", "MIX"):
                            sfe_tabulka.append(naceneny_item)
                        elif program == "BFE":
                            bfe_tabulka.append(naceneny_item)
                        elif program not in ("SFE", "MIX", "BFE"):
                            neznamy_program_tabulka.append(naceneny_item)
                        uz_nacenene_vrcholy.append(vrchol)
                        if vrchol in zatim_nenacenitelne_itemy_tabulka:
                            zatim_nenacenitelne_itemy_tabulka.remove(vrchol)    
                # 2. Pokud se jedna o item pod platnym kalkulacnim projektem → nacenit do tabulky
                    elif vrchol[0:9] in platne_kalkulacni_projekty:
                        # print(f'vyrabena pod platnym kalkulakem {vrchol}')
                        program = funkce_prace.dotaz_pn_program(vrchol[9:len(vrchol)], databaze_pro_dotaz_programy, kvp_programy)[0].split(":")[1]    
                        # print(f'vrchol {vrchol} - {program}')
                        naceneny_item, bom_qty_check = KOMBO_funkce_kusovnik.naceneni_do_tabulek(vrchol, parametry, program, kusovnik, kusovnik_bom_qty)
                        # zachyceni vyrabenych itemu s 0 bom qty / PLACARDU s kusovnikem a pridani do chyb.
                        for item in bom_qty_check[1]:
                            if item not in vsechny_vyrabene_bom_qty_0:
                                vsechny_vyrabene_bom_qty_0.append(item)

                        if program in ("SFE", "MIX"):
                            sfe_tabulka.append(naceneny_item)
                        elif program == "BFE":
                            bfe_tabulka.append(naceneny_item)
                        elif program not in ("SFE", "MIX", "BFE"):
                            neznamy_program_tabulka.append(naceneny_item)
                        uz_nacenene_vrcholy.append(vrchol[9:len(vrchol)])
                        if vrchol[9:len(vrchol)] in zatim_nenacenitelne_itemy_tabulka:
                            zatim_nenacenitelne_itemy_tabulka.remove(vrchol[9:len(vrchol)])                         
                # 3. Pokud se jedna o MAN PLACARD / ID PLACARD → naceni do tabulky           
                    elif (KOMBO_funkce_kusovnik.je_to_man_placard(vrchol, parametry) or KOMBO_funkce_kusovnik.je_to_id_placard(vrchol, parametry)):
                        if (vrchol[0:3] != "PMP" and vrchol not in uz_nacenene_vrcholy) or (vrchol[0:3] == "PMP" and vrchol[9:len(vrchol)] not in uz_nacenene_vrcholy):                            
                            # print(f'MAN PLACARD / ID PLACARD {vrchol}')
                            # Pokud se jedna o PLACARD pod PMP projektem
                            if vrchol[0:3] == "PMP":
                                program = funkce_prace.dotaz_pn_program(vrchol[9:len(vrchol)], databaze_pro_dotaz_programy, kvp_programy)[0].split(":")[1]
                            else:                           
                                program = funkce_prace.dotaz_pn_program(vrchol, databaze_pro_dotaz_programy, kvp_programy)[0].split(":")[1]
                            # print(f'vrchol {vrchol} - {program}')
                            naceneny_item, bom_qty_check = KOMBO_funkce_kusovnik.naceneni_do_tabulek(vrchol, parametry, program, kusovnik, kusovnik_bom_qty)
                            # zachyceni vyrabenych itemu s 0 bom qty / PLACARDU s kusovnikem a pridani do chyb.
                            for item in bom_qty_check[1]:
                                if item not in vsechny_vyrabene_bom_qty_0:
                                    vsechny_vyrabene_bom_qty_0.append(item)

                            if program in ("SFE", "MIX"):
                                sfe_tabulka.append(naceneny_item)
                            elif program == "BFE":
                                bfe_tabulka.append(naceneny_item)
                            elif program not in ("SFE", "MIX", "BFE"):
                                neznamy_program_tabulka.append(naceneny_item)
                            if vrchol[0:3] != "PMP":
                                uz_nacenene_vrcholy.append(vrchol)
                                if vrchol in zatim_nenacenitelne_itemy_tabulka:
                                    zatim_nenacenitelne_itemy_tabulka.remove(vrchol)                                 
                            elif vrchol[0:3] == "PMP":
                                uz_nacenene_vrcholy.append(vrchol[9:len(vrchol)])
                                if vrchol[9:len(vrchol)] in zatim_nenacenitelne_itemy_tabulka:
                                    zatim_nenacenitelne_itemy_tabulka.remove(vrchol[9:len(vrchol)])                                  
                    else:
                        if vrchol[0:3] == "PMP":
                            if vrchol[9:len(vrchol)] not in uz_nacenene_vrcholy:
                                zatim_nenacenitelne_itemy_tabulka.append(vrchol)
                        else:                        
                            if vrchol not in uz_nacenene_vrcholy:
                                zatim_nenacenitelne_itemy_tabulka.append(vrchol)

                if len(vrchol_chyby) != 0:
                    print(f'Vrchol {vrchol} chyby {vrchol_chyby}')

                for linka in vrchol_k_zaplanovani:
                    output_file.write(str(linka))
                    output_file.write('\n')
                
                # KOMBO_funkce_kusovnik.zaplanovani_do_excelu(vse_k_zaplanovani)    

output_file.close()

## Vytisknuti vysledku a chyb na konci programu.
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
        elif parametry.get(item).get("supplier") == "I00000008":    
            print(f'Lamphun_A:{item}:{parametry.get(item).get("description")}')
        else:
            print(f'anonymni_:{item}:{parametry.get(item).get("description")}')

if len(vsechny_neplatne_routingy) != 0:
    print("\nSeznam vsech dilu s neplatnym routingem z reportu:")
    for item in vsechny_neplatne_routingy:
        if item[0:3] == "PMP":
            print(item[0:9]+":"+str(item[9:len(item)]))
        elif parametry.get(item).get("supplier") == "I00000008":    
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
        elif parametry.get(item).get("supplier") == "I00000008":    
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
        elif parametry.get(item).get("supplier") == "I00000008":    
            print("Lamphun_A:"+str(item))
        else:
            print("anonymni_:"+str(item))

if len(vsechny_vyrabene_bom_qty_0) != 0:
    print("\nSeznam Manufactured dilu s 0 BOM qtu / PLACARDu s nesmazanym kusovnikem:")
    print(vsechny_vyrabene_bom_qty_0)
    for item in vsechny_vyrabene_bom_qty_0:
        if item[0][0:3] == "PMP":
            print(f'{item[0][0:9]}:{item[0][9:len(item[0])]}:{parametry.get(item[0]).get("description")} v lince {item[1]}')
        elif parametry.get(item[0]).get("supplier") == "I00000008":    
            print(f'Lamphun_A:{item[0]}:{parametry.get(item[0]).get("description")} v lince {item[1]}')
        else:
            print(f'anonymni_:{item[0]}:{parametry.get(item[0]).get("description")} v lince {item[1]}')            

if len(vsechny_manufactured_polozky_pod_projekt) != 0:
    print("\nSeznam Manufactured dilu PREPNOUT POD PROJEKT:")
    for item in vsechny_manufactured_polozky_pod_projekt:
        print(item)

print("\n")
print("Vysledek pro naceneni:\n")
print("SFE tabulka:\n")
for naceneny_vrchol in sfe_tabulka:
    print(naceneny_vrchol)
print("\n")    
print("BFE tabulka:\n")
for naceneny_vrchol in bfe_tabulka:
    print(naceneny_vrchol)
print("\n") 
print("Neznamy program tabulka:\n")
for naceneny_vrchol in neznamy_program_tabulka:
    print(naceneny_vrchol)
print("\n")

zatim_nenacenitelne_itemy_tabulka2 = []
for error in zatim_nenacenitelne_itemy_tabulka:
    # print(error)
    if error[0:3] == "PMP":
        if error[9:len(error)] not in uz_nacenene_vrcholy:
            zatim_nenacenitelne_itemy_tabulka2.append(error)
    elif error not in uz_nacenene_vrcholy:
            zatim_nenacenitelne_itemy_tabulka2.append(error)
zatim_nenacenitelne_itemy_tabulka = zatim_nenacenitelne_itemy_tabulka2

if len(zatim_nenacenitelne_itemy_tabulka) != 0:
    print("Zatim nenacenitelne polozky tabulka:\n(Nejedna se o\n1) ciste nakupovanou s platnymi PUR daty.\n2) vyrabenou polozku pod platnym kalkulacnim projektem.\n3) MAN PLACARD / ID PLACARD.)\n")
    for error_vrchol in zatim_nenacenitelne_itemy_tabulka:
        print(error_vrchol)

# print(parametry)    
print("\nKONEC PROGRAMU")
