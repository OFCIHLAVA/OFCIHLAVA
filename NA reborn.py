import datetime

from Programy import funkce_prace

### Pouzivane funkce

def file_present_check(path_to_file):
    import os, sys
    
    try:
        if os.stat(path_to_file).st_size == 0:
            print(f'\nPOZOR! V souboru {path_to_file} Nejsou zadna data. Neni co zpracovavat . . .')
            input(f'Stiskni ENTER pro ukonceni programu')
            sys.exit()
    except FileNotFoundError:
        print(f'\nPOZOR! nebyl nalezen soubor {path_to_file}. Neni co zpracovavat . . .')    
        input(f'Stiskni ENTER pro ukonceni programu')
        sys.exit()

def empty_value(data_field): # Overi, zda se jedna o prazdnou hodnotu.
    return True if str(data_field) == "NULL" else False    

def valid_exdate(exdate=datetime): # Overi, zda ma dil platne expiry date.
    import datetime
    return True if exdate > datetime.date.today() else False 

def valid_routing(routing=int): # Overi podle delky routingu, jestli se jedna o neplatny routing.
    return True if routing <= 2000 else False  # (ve vstupnich datech je kazdy neplatny routing >2000 dni)

def projektovy_dil(part_number=str): # Urceni, zda se jedna o projektovy dil.
    return True if part_number[0:3] == "PMP" else False

def anonymni_dil_projektu(pmp_part_number=str): # Ziskano, anonymniho cisla projektoveho dilu.
    ppn = pmp_part_number
    anonym = ppn[9:len(ppn)] # Projektovy dil ma format: PMPxxxxxxANONYMNIDIL (xxxxxx je cislo projektu).
    return anonym

def data_import(cq_export_cesta): # Nacteni dat z CQ reportu.
    with open(cq_export_cesta,'r', encoding='Windows-1250') as input_file:
        data = input_file.readlines()
        input_file.close()
    return data

def data_headings(data): # Vytvoreni zahlavi sloupcu z reportu.
    for line in data:
        if "***" in line[2:5]: # Identifikator linky se zahlavim z cq reportu
            data_headings = [pole.strip() for pole in line.split("|")]
            del data_headings[0]
            break
    return data_headings  

def import_data_cleaning(data): # Ocisteni dat a nahrazeni prazdnych poli za "0".
    cl_data = list()    
    for line in data: # Ocisteni dat a 
        if line[0] == "|":
            linka = [pole.strip() for pole in line.split("|")] # Ocisteni dat.
            for i, pole in enumerate(linka): # nahrazeni praydnych poli za "0".
                if len(pole) == 0:
                    linka[i] = "NULL"
            # smazani prvniho pole(vznikne s hodnotou 0 pri rozdeleni linky).
            del linka[0]
            cl_data.append(linka)
            data = cl_data
    return data

def data_date_formating(data, data_headings): # Prevedeni pole datumu na date format.
    
    import datetime

    datumy_indexy = [data_headings.index(heading) for heading in data_headings if "DATE" in heading.upper()]
    for line in data: # Prevedeni pole datumu na date format.
        for datum_index in datumy_indexy:      
            # Pokud je dal prazdny datum, uz to byl posledni item na lince → preskocit na dalsi.
            if empty_value(line[datum_index]):
                break
            den, mesic, rok = line[datum_index].split("/")
            datum = datetime.date(int(rok), int(mesic), int(den))
            line[datum_index] = datum
    return data

def delete_lvl_from_heading(heading=str):
    # charaktery k odstraneni z nazvu sloupcu, aby se odstranila informace o levelu. (Pokud by se upravil nazev v zahlavi, je nutno tady take upravit.)
    to_delete = [
        "vrchol",
        "lvl",
        "1",
        "2",
        "3",
        "4",
        "5",
        "6",
        "7",
        "8",
        "9",
        "10",
        "11"
        ]
    for char in to_delete:
        heading = heading.replace(char,"")
    heading = heading.strip()
    return heading

def items_database(data, headings): # Vytvoreni dictionary databaze jednotlivych itemu s jejich linkamy a daty z reportu.
    database_dictionary = dict() # Vsechny vrcholy z reportu a jejich linky priprava dict.
    # Item indexy
    part_number_indexy = [headings.index(heading) for heading in headings if "Part number".upper() in heading.upper()]
    # print(part_number_indexy)

    # Upravene zahlavi bez informace o levelech itemu.
    no_lvls_headings = [delete_lvl_from_heading(heading) for heading in headings]
    # print(no_lvls_headings)

    # 1. Projizdet linky a itemy v nich.
    for i, line in enumerate(data):        
        # print(f'Linka {i}')
        # Projet kazdy item v lince kusovniku
        for pn_index in part_number_indexy: 
            part_number = line[pn_index]
            # print(f'{part_number} na pozici {pn_index}')
            
            # Pokud je item NULL (prazdna hodnota), uz to byl posledni item v lince → poskocit na dalsi linku.
            if empty_value(part_number):
                # print(f'Toto byl posledni item v lince → preskakuji na dalsi linku.')
                break
            # Jinak, je PN platny item a dal proveruji, jestli ho pridavat:
            
            # 2. Pokud pn jeste neni v parametrech:
            if part_number not in database_dictionary:
                # Overit, zda ma reseny item platne exdate a chceme ho potencialne pridavat do databaze
                # Dohledani exdate k resenemu itemu.]
                for heading in no_lvls_headings[pn_index:]:
                    # BERE SE PRVNI NALEZENE EXDATE VPRAVO OD ZKOUMANEHO ITEMU.
                    if "EXDATE" in heading.upper():
                        exdate_index = pn_index + no_lvls_headings[pn_index:].index(heading)
                        exdate = line[exdate_index]
                        # print(exdate)
                        break # Az se najde prvni exdate (k momentalnimu pn) → hned break → POKRACOVANI NA DALSI ITEM V LINCE.
                # 3. Pokud item nema platny EX. date → nepridavat ho ani zandy item v lince za nim → preskocit na dalsi linku.
                if not valid_exdate(exdate):
                    # print(f'neplatne exdate')
                    break
                # 4. Jinak Pridat data reseneho itemu da databaze.
                # Dohledani odpovidajicich dat k itemu v resene lince.
                
                # Dalsi item v lince za resenym itemem.
                next_pn_index = part_number_indexy[min(part_number_indexy.index(pn_index)+1, len(part_number_indexy)-1)] # Osetreno proti out of bounds index, kdyz je to posledni item uz.    
        
                # Data jen pro reseny item:
                if pn_index != next_pn_index: # Pokud se NEJEDNA o posledni item v lince, jeho data se bere linka od nej k dalsimu itemu. 
                    pn_data_line = line[pn_index:next_pn_index] 
                else: # Pokud se JEDNA o posledni item v lince, jeho data se bere linka od nej az do konce linky. 
                    pn_data_line = line[pn_index:len(line)]

                # Sestaveni dat linky jako dict pro ulozeni do databaze.                   
                data_dict = {} # Samotna data na linkach.
                for i, data_field in enumerate(pn_data_line):
                    data_dict[no_lvls_headings[pn_index+i]] = data_field # Sestaveni dat z linky jako dict s jmeny sloupcu zahlavi ocistenych o lvl informace reportu jako klic.                
                # Pridat do databaze
                database_dictionary[part_number] = data_dict # Sestaveni linek predchoziho vrcholu jako dict vrchol : jeho linky s daty viz. vyse. 

                # 5. Pokud je reseny dil projektovy dil, updatovat anonymni polozku o routing z projektoveho, pokud je platny.    
 
                # Pokud se jedna o projektovy dil.
                if projektovy_dil(part_number):
                    # Pokud k nemu existuje anonymni polozka.
                    if anonymni_dil_projektu(part_number) in database_dictionary:
                        # Kouknout jestli ma pr. dil platny routing a STD rouding a pokud ANO → pridat updatovat tento routing pro anonymni polozku.                        
                        # PMP routing 
                        projektovy_routing = database_dictionary.get(part_number).get("routing LT")
                        # print(f'p routing: {projektovy_routing}, {type(projektovy_routing)}')
                        if not empty_value(projektovy_routing):
                            projektovy_routing = int(projektovy_routing)
                            if valid_routing(projektovy_routing):
                                database_dictionary[part_number]["routing LT"] = projektovy_routing
                        # STD routing 
                        std_projektovy_routing = database_dictionary.get(part_number).get("STD routing LT")
                        # print(f'std routing: {std_projektovy_routing}, {type(std_projektovy_routing)}')
                        if not empty_value(std_projektovy_routing):
                            std_projektovy_routing = int(std_projektovy_routing)
                            if valid_routing(std_projektovy_routing):
                                database_dictionary[part_number]["STD routing LT"] = std_projektovy_routing                         
                # 6. Pokud je to posledni item na lince → preskok na dalsi linku.
                if pn_index == next_pn_index:
                    # print(f'Toto byl posledni item v lince → preskakuji na dalsi linku.')
                    break
            # 7. Pokud pn uz je v databazi.
            #  Pokud je to posledni item na lince → preskok na dalsi linku.
            if pn_index == next_pn_index:
                # print(f'Toto byl posledni item v lince → preskakuji na dalsi linku.')
                break
    return database_dictionary

def effective_bom_and_items_database(data, headings): # Vytvoreni effective kussovniku lineka a dictionary databaze jednotlivych itemu z reportu.
    items_database_dictionary = dict() # Vsechny vrcholy z reportu a jejich linky priprava dict.
    effective_bom_lines = list()
    # Item indexy
    part_number_indexy = [headings.index(heading) for heading in headings if "Part number".upper() in heading.upper()]
    # print(part_number_indexy)

    # Upravene zahlavi bez informace o levelech itemu.
    no_lvls_headings = [delete_lvl_from_heading(heading) for heading in headings]
    # print(no_lvls_headings)

    # 1. Projizdet linky a itemy v nich.
    for i, line in enumerate(data):        
        
        effective_bom_line = list()
        # print(f'Linka {i}')
        # Projet kazdy item v lince kusovniku
        for pn_index in part_number_indexy: 
            part_number = line[pn_index]
            # print(f'{part_number} na pozici {pn_index}')
            
            # Pokud je item NULL (prazdna hodnota), uz to byl posledni item v lince → poskocit na dalsi linku.
            if empty_value(part_number):
                # print(f'Toto byl posledni item v lince → preskakuji na dalsi linku.')
                break
            # Jinak, je PN platny item a dal proveruji, jestli ho pridavat:
            

            # 2. Overit, zda ma reseny item platne exdate a chceme ho potencialne pridavat do databaze
            # Dohledani exdate k resenemu itemu.]
            for heading in no_lvls_headings[pn_index:]:
                # BERE SE PRVNI NALEZENE EXDATE VPRAVO OD ZKOUMANEHO ITEMU.
                if "EXDATE" in heading.upper():
                    exdate_index = pn_index + no_lvls_headings[pn_index:].index(heading)
                    exdate = line[exdate_index]
                    # print(exdate)
                    break # Az se najde prvni exdate (k momentalnimu pn) → hned break → POKRACOVANI NA DALSI ITEM V LINCE.
            #  Pokud item nema platny EX. date → nepridavat ho ani zandy item v lince za nim → preskocit na dalsi linku.
            if not valid_exdate(exdate):
                # print(f'neplatne exdate')
                break            
            # 3. Jinak Pridat data reseneho itemu da databaze a effective kusovniku.
            # Dohledani odpovidajicich dat k itemu v resene lince.           
               
            # Dalsi item v lince za resenym itemem.
            next_pn_index = part_number_indexy[min(part_number_indexy.index(pn_index)+1, len(part_number_indexy)-1)] # Osetreno proti out of bounds index, kdyz je to posledni item uz.    
    
            # Data jen pro reseny item:
            if pn_index != next_pn_index: # Pokud se NEJEDNA o posledni item v lince, jeho data se bere linka od nej k dalsimu itemu. 
                pn_data_line = line[pn_index:next_pn_index] 
            else: # Pokud se JEDNA o posledni item v lince, jeho data se bere linka od nej az do konce linky. 
                pn_data_line = line[pn_index:len(line)]
            # Sestaveni dat linky jako dict pro ulozeni do databaze.                   
            data_dict = {} # Samotna data na linkach.
            for i, data_field in enumerate(pn_data_line):
                data_dict[no_lvls_headings[pn_index+i]] = data_field # Sestaveni dat z linky jako dict s jmeny sloupcu zahlavi ocistenych o lvl informace reportu jako klic.                
            
            # Pridani do effective kusovniku.
            bom_qty = data_dict.get("BOM qty")
            warehouse = data_dict.get("WHS")
            effective_bom_line.append([part_number, {"BOM qty":bom_qty}, {"WHS":warehouse}])
             
            # Pridat do databaze
            # 2. Pokud pn jeste neni v parametrech:
            if part_number not in items_database_dictionary: 
                items_database_dictionary[part_number] = data_dict # Sestaveni linek predchoziho vrcholu jako dict vrchol : jeho linky s daty viz. vyse. 

                # 5. Pokud je reseny dil projektovy dil, updatovat anonymni polozku o routing z projektoveho, pokud je platny.    
 
                # Pokud se jedna o projektovy dil.
                if projektovy_dil(part_number):
                    # Pokud k nemu existuje anonymni polozka.
                    if anonymni_dil_projektu(part_number) in items_database_dictionary:
                        # Kouknout jestli ma pr. dil platny routing a STD rouding a pokud ANO → pridat updatovat tento routing pro anonymni polozku.                        
                        # PMP routing 
                        projektovy_routing = items_database_dictionary.get(part_number).get("routing LT")
                        # print(f'p routing: {projektovy_routing}, {type(projektovy_routing)}')
                        if not empty_value(projektovy_routing):
                            projektovy_routing = int(projektovy_routing)
                            if valid_routing(projektovy_routing):
                                items_database_dictionary[part_number]["routing LT"] = projektovy_routing
                        # STD routing 
                        std_projektovy_routing = items_database_dictionary.get(part_number).get("STD routing LT")
                        # print(f'std routing: {std_projektovy_routing}, {type(std_projektovy_routing)}')
                        if not empty_value(std_projektovy_routing):
                            std_projektovy_routing = int(std_projektovy_routing)
                            if valid_routing(std_projektovy_routing):
                                items_database_dictionary[part_number]["STD routing LT"] = std_projektovy_routing                         
                # 6. Pokud je to posledni item na lince → preskok na dalsi linku.
            if pn_index == next_pn_index:
                # print(f'Toto byl posledni item v lince → preskakuji na dalsi linku.')
                break
            # 7. Pokud pn uz je v databazi. MOZNA RESI? NECHAT ZATIM OTEVRENO.
            # else:
            #     print(f'{part_number} uz je')
            #     if valid_exdate(exdate): # Pokud ma platne exdate:
            #         # Zkontrolovat stavajici polozku, jestli ma nejake NUll hodnoty.
            #         for data_nazev, hodnota in database_dictionary.get(part_number).items():
            #             if hodnota == "NULL"
            #             print(data_nazev, hodnota)
        effective_bom_lines.append(effective_bom_line)
    return effective_bom_lines, items_database_dictionary

def struktura_kusovniku(bom_data_lines=list): # Sestaveni struktury effective kusovniku po levelech.
    
    kusovnik_struktura = dict()
    # Po linkach projizdi kusovnik a kouka na jednotlive itemy v lince, jestli uz jsou ve strukture zarazene pod svym naddilem z dane linky, pripadne je k nemu doplni. 
    for linka in bom_data_lines:
        # print(f'LINKA: {linka}')
        for level, item_bomqty_whs in enumerate(linka):
            # print(f'LVL: {level}')

            pn = item_bomqty_whs[0]
            bom_qty = item_bomqty_whs[1].get("BOM qty")
            warehouse = item_bomqty_whs[2].get("WHS")
            # print(f'Item: {pn}, Qty: {bom_qty}, WHS: {warehouse}')            
            # Vrchol
            if level == 0:
                if pn not in kusovnik_struktura:
                    kusovnik_struktura[pn] = {"BOM qty":bom_qty, "WHS":warehouse, "poddily":dict()}
                    # print(kusovnik_struktura)
            elif level == 1:           
                if pn not in kusovnik_struktura.get(linka[0][0]).get("poddily"):
                    kusovnik_struktura[linka[0][0]]["poddily"][pn] = {"BOM qty":bom_qty, "WHS":warehouse, "poddily":dict()}
                    # print(kusovnik_struktura)
            elif level == 2:
                if pn not in kusovnik_struktura.get(linka[0][0]).get("poddily").get(linka[1][0]).get("poddily"):
                    kusovnik_struktura[linka[0][0]]["poddily"][linka[1][0]]["poddily"][pn] = {"BOM qty":bom_qty, "WHS":warehouse, "poddily":dict()}
                    # print(kusovnik_struktura)           
            elif level == 3:
                if pn not in kusovnik_struktura.get(linka[0][0]).get("poddily").get(linka[1][0]).get("poddily").get(linka[2][0]).get("poddily"):
                    kusovnik_struktura[linka[0][0]]["poddily"][linka[1][0]]["poddily"][linka[2][0]]["poddily"][pn] = {"BOM qty":bom_qty, "WHS":warehouse, "poddily":dict()}
                    # print(kusovnik_struktura) 
            elif level == 4:
                if pn not in kusovnik_struktura.get(linka[0][0]).get("poddily").get(linka[1][0]).get("poddily").get(linka[2][0]).get("poddily").get(linka[3][0]).get("poddily"):
                    kusovnik_struktura[linka[0][0]]["poddily"][linka[1][0]]["poddily"][linka[2][0]]["poddily"][linka[3][0]]["poddily"][pn] = {"BOM qty":bom_qty, "WHS":warehouse, "poddily":dict()}
                    # print(kusovnik_struktura)
            elif level == 5:
                if pn not in kusovnik_struktura.get(linka[0][0]).get("poddily").get(linka[1][0]).get("poddily").get(linka[2][0]).get("poddily").get(linka[3][0]).get("poddily").get(linka[4][0]).get("poddily"):
                    kusovnik_struktura[linka[0][0]]["poddily"][linka[1][0]]["poddily"][linka[2][0]]["poddily"][linka[3][0]]["poddily"][linka[4][0]]["poddily"][pn] = {"BOM qty":bom_qty, "WHS":warehouse, "poddily":dict()}
                    # print(kusovnik_struktura)
            elif level == 6:
                if pn not in kusovnik_struktura.get(linka[0][0]).get("poddily").get(linka[1][0]).get("poddily").get(linka[2][0]).get("poddily").get(linka[3][0]).get("poddily").get(linka[4][0]).get("poddily").get(linka[5][0]).get("poddily"):
                    kusovnik_struktura[linka[0][0]]["poddily"][linka[1][0]]["poddily"][linka[2][0]]["poddily"][linka[3][0]]["poddily"][linka[4][0]]["poddily"][linka[5][0]]["poddily"][pn] = {"BOM qty":bom_qty, "WHS":warehouse, "poddily":dict()}
                    # print(kusovnik_struktura)
            elif level == 7:
                if pn not in kusovnik_struktura.get(linka[0][0]).get("poddily").get(linka[1][0]).get("poddily").get(linka[2][0]).get("poddily").get(linka[3][0]).get("poddily").get(linka[4][0]).get("poddily").get(linka[5][0]).get("poddily").get(linka[6][0]).get("poddily"):
                    kusovnik_struktura[linka[0][0]]["poddily"][linka[1][0]]["poddily"][linka[2][0]]["poddily"][linka[3][0]]["poddily"][linka[4][0]]["poddily"][linka[5][0]]["poddily"][linka[6][0]]["poddily"][pn] = {"BOM qty":bom_qty, "WHS":warehouse, "poddily":dict()}
                    # print(kusovnik_struktura)
            elif level == 8:
                if pn not in kusovnik_struktura.get(linka[0][0]).get("poddily").get(linka[1][0]).get("poddily").get(linka[2][0]).get("poddily").get(linka[3][0]).get("poddily").get(linka[4][0]).get("poddily").get(linka[5][0]).get("poddily").get(linka[6][0]).get("poddily").get(linka[7][0]).get("poddily"):
                    kusovnik_struktura[linka[0][0]]["poddily"][linka[1][0]]["poddily"][linka[2][0]]["poddily"][linka[3][0]]["poddily"][linka[4][0]]["poddily"][linka[5][0]]["poddily"][linka[6][0]]["poddily"][linka[7][0]]["poddily"][pn] = {"BOM qty":bom_qty, "WHS":warehouse, "poddily":dict()}
                    # print(kusovnik_struktura)
            elif level == 9:
                if pn not in kusovnik_struktura.get(linka[0][0]).get("poddily").get(linka[1][0]).get("poddily").get(linka[2][0]).get("poddily").get(linka[3][0]).get("poddily").get(linka[4][0]).get("poddily").get(linka[5][0]).get("poddily").get(linka[6][0]).get("poddily").get(linka[7][0]).get("poddily").get(linka[8][0]).get("poddily"):
                    kusovnik_struktura[linka[0][0]]["poddily"][linka[1][0]]["poddily"][linka[2][0]]["poddily"][linka[3][0]]["poddily"][linka[4][0]]["poddily"][linka[5][0]]["poddily"][linka[6][0]]["poddily"][linka[7][0]]["poddily"][linka[8][0]]["poddily"][pn] = {"BOM qty":bom_qty, "WHS":warehouse, "poddily":dict()}
                    # print(kusovnik_struktura)
            elif level == 10:
                if pn not in kusovnik_struktura.get(linka[0][0]).get("poddily").get(linka[1][0]).get("poddily").get(linka[2][0]).get("poddily").get(linka[3][0]).get("poddily").get(linka[4][0]).get("poddily").get(linka[5][0]).get("poddily").get(linka[6][0]).get("poddily").get(linka[7][0]).get("poddily").get(linka[8][0]).get("poddily").get(linka[9][0]).get("poddily"):
                    kusovnik_struktura[linka[0][0]]["poddily"][linka[1][0]]["poddily"][linka[2][0]]["poddily"][linka[3][0]]["poddily"][linka[4][0]]["poddily"][linka[5][0]]["poddily"][linka[6][0]]["poddily"][linka[7][0]]["poddily"][linka[8][0]]["poddily"][linka[9][0]]["poddily"][pn] = {"BOM qty":bom_qty, "WHS":warehouse, "poddily":dict()}
    return kusovnik_struktura

def print_struktura_kusovniku(kusovnik, hloubka_kusovniku): # VYTISKNE STRUKTURU KUSOVNIKU S JEHO LEVELY NA ZAKLADE DICT STRUKTRURY KUSOVNIKU.
 
    # print(f'Hloubka pred spustenim = {hloubka_kusovniku}')
    hloubka_kusovniku += 1
    # print(f'Hloubka pred spustenim = {hloubka_kusovniku}')
    for dil, qty_whs_a_poddily in kusovnik.items(): 
        # print(f'Dil: {dil}')        
        # print(f'Dil bez projektu: {dil_bez_projektu}')
        # print(f'{"--"* hloubka_kusovniku}Level_{hloubka_kusovniku}: {dil}')
        print(f'{"  "*hloubka_kusovniku}Level_{hloubka_kusovniku}: {dil}, BOM Qty: {kusovnik[dil]["BOM qty"]}, Warehouse: {kusovnik[dil]["WHS"]}')
        # result_to_print.append(f'Level_{hloubka_kusovniku}: {dil}, BOM Qty: {kusovnik[dil]["BOM qty"]}')
        # print(result_to_print)
        for key, values in qty_whs_a_poddily.items():
            if key == "poddily":
                print_struktura_kusovniku(values, hloubka_kusovniku)

### Pouzite soubory:
# CQ bom export
bom_export = "Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\1_LT + Nastavení položek\\LT LN export\\report.txt"
# txt databaze pro SFExBFExMIX
path_kusovniky_databaze = 'Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\5_SFExBFExMIX\\databaze\\databaze boud s kusovniky.txt'
path_program_databaze = 'Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\5_SFExBFExMIX\\databaze\\seznam programu.txt'




### Kontorola spravnosti cest k pouzivanym souborum a jestli jsou v nich data.

file_present_check(bom_export)
file_present_check(path_kusovniky_databaze)
file_present_check(path_program_databaze)



### Nacteni dat a ocisteni dat.

# Import
bom_import = data_import(bom_export)
# Zahlavi
zahlavi_bom_dat = data_headings(bom_import)
# Ocisgteni dat + oznaceni prazdnych poli jako NULL
bom_data_no_headings = import_data_cleaning(bom_import)
# Upraveni datumu na platny format
bom_data_no_headings = data_date_formating(bom_data_no_headings, zahlavi_bom_dat)



### Priprava Modulu na zjistovani SFExBFExMIX

databaze_pro_dotaz_programy = funkce_prace.nacteni_databaze_boud_pro_dotaz(path_kusovniky_databaze)
kvp_programy = funkce_prace.programy_boud(path_program_databaze)



### Sestaveni databaze atributu platnych itemu.

effective_bom_lines, databaze_itemu = effective_bom_and_items_database(bom_data_no_headings, zahlavi_bom_dat)

# Kontrolni tisk deffective kusovniku itemu
# for line in effective_bom_lines:
# 
#     only_pn_line = []
#     print(line)
#     for pn_qty in line:
#         for pn, qty in pn_qty.items():
#             only_pn_line.append(pn)
#     print(only_pn_line)

# Kontrolni tisk databaze itemu
# for pn, pn_data in databaze_itemu.items():
#     print(pn)
#     for data_field_name, data in pn_data.items():
#         print(f'    {data_field_name}: {data}')



### Sestaveni struktury effective kusovniku s jeho levely.
effective_bom_structure = struktura_kusovniku(effective_bom_lines)



### Vytisteni struktury effective kusovniku s jeho levely.
print_struktura_kusovniku(effective_bom_structure, 0)