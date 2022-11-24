import datetime
from .excel_data import dnesni_datum, neplatne_datum_ln


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

def data_headings_order_plan(op_data): # Vytvoreni zahlavi sloupcu z reportu.
    for line in op_data:
        if "***" in line[2:5]: # Identifikator linky se zahlavim z cq reportu.
            data_headings = [pole.strip() for pole in line.split("|")]
            data_headings[0] = "Item line in LN" # Prepsani "***" v prvnim sloupci, kde bude pozdeji doplneno poradi linky v cq reportu.
            break
    return data_headings

def data_headings_master_plan(mp_data): # Vytvoreni zahlavi sloupcu z reportu.
    for line in mp_data:
        if "***" in line[2:5]: # Identifikator linky se zahlavim z cq reportu.
            data_headings = [pole.strip() for pole in line.split("|")]
            data_headings[0] = "Row in Master plan" # Prepsani "***" v prvnim sloupci, kde bude pozdeji doplneno poradi linky v Master plan reportu.
            data_headings.append("Availability") # Pridani heading Availibility na konec - bude doplneno pozdeji.
            break
    return data_headings    

def import_data_cleaning(data): # Ocisteni dat a nahrazeni prazdnych poli za "0".
    cl_data = list()    
    for line in data: # Ocisteni dat a 
        if line[0] == "|":
            linka = [pole.strip() for pole in line.split("|")] # Ocisteni dat.
            # print(linka)
            for i, pole in enumerate(linka): # nahrazeni praydnych poli za "0".
                if len(pole) == 0:
                    linka[i] = "0"
            # smazani prvniho pole(vznikne s hodnotou 0 pri rozdeleni linky).
            del linka[0]
            cl_data.append(linka)
            data = cl_data
    return data

def add_line_id_to_order_plan_data(op_data, op_data_headings): # Pridani ID linky pro jednotlive vrcholy do cq dat jako int na pozici 0 kazde linky. Zvlast resen sklad PZN100 a PZN105.
    pocet_linek_vrcholu_PZN100 = {}
    pocet_linek_vrcholu_PZN105 = {}
    pocet_linek_vrcholu_PZN310 = {}
    
    for heading in op_data_headings: # Nalezeni indexu sloupcu skladu a vrcholu v cq importu.
        if "CLUSTER" in heading.upper() or "WAREHOUSE" in heading.upper():
            sklad_index = op_data_headings.index(heading)
        if heading.upper() == "ITEM":
            vrchol_index = op_data_headings.index(heading)
    
    for line in op_data:
        sklad = line[sklad_index-1] # -1 Protoze v headings uz je zahlavi 1. skoupce (id line), ale v datech jeste ne, teprve ho tam ted pridame. (data jsou kratsi o 1 pozici).
        vrchol = line[vrchol_index-1] # -1 Protoze v headings uz je zahlavi 1. skoupce (id line), ale v datech jeste ne, teprve ho tam ted pridame. (data jsou kratsi o 1 pozici).

        #print(sklad)
        #print(vrchol)

        # PZN100
        if sklad.upper() == "PZN100":
            if vrchol not in pocet_linek_vrcholu_PZN100:
                pocet_linek_vrcholu_PZN100[vrchol] = 0 
            line.insert(0,pocet_linek_vrcholu_PZN100.get(vrchol)+1) # Vlozeni sloupce ID linky pred stavavjici data. (rozsireni dat o jeden sloupec)
            # print(line)
            pocet_linek_vrcholu_PZN100[vrchol] +=1

        # PZN105
        elif sklad.upper() == "PZ5" or sklad.upper() == "PZN105":
            # print(f"PZN105")
            if vrchol not in pocet_linek_vrcholu_PZN105:
                pocet_linek_vrcholu_PZN105[vrchol] = 0 
            line.insert(0,pocet_linek_vrcholu_PZN105.get(vrchol)+1)
            pocet_linek_vrcholu_PZN105[vrchol] +=1

        # PZN310
        elif sklad.upper() == "PZ4" or sklad.upper() == "PZN310":
            # print(f"PZN310")
            if vrchol not in pocet_linek_vrcholu_PZN310:
                pocet_linek_vrcholu_PZN310[vrchol] = 0 
            line.insert(0,pocet_linek_vrcholu_PZN310.get(vrchol)+1)
            pocet_linek_vrcholu_PZN310[vrchol] +=1
    return op_data

def add_row_number_to_master_plan_data(mp_data): # Pridani poradi linky v Master planu do cq data Master planu.  
    for row, linka in enumerate(mp_data):
        linka.insert(0,row+1)
    return(mp_data)

def add_availability_master_plan_data(mp_data, mp_data_headings): # Pridani sloupce Availability do cq data Master planu.  
    for heading in mp_data_headings:
        if heading.upper() == "INVENTORY ON HAND":
            inventory_on_hand_index = mp_data_headings.index(heading)
        elif heading.upper() == "KLIC1":
            klic1_index = mp_data_headings.index(heading)
        elif heading.upper() == "ORDERED QTY":
            ordered_qty_index = mp_data_headings.index(heading)

    item_already_requested_qty_before_current_line = dict()

    for line in mp_data:
        inventory_on_hand = float(line[inventory_on_hand_index].replace(",",""))
        ordered_qty = float(line[ordered_qty_index].replace(",",""))
        item = line[klic1_index]

        # pokud jeste ten item nebyl req before, pridej ho do uz requested s ordered qty na lince:
        if not item_already_requested_qty_before_current_line.get(item):
            item_already_requested_qty_before_current_line[item] = ordered_qty
        # Pokud uz tam byl, pricti k te hodnote co uz tam byla ordered qty na lince:
        else:
            item_already_requested_qty_before_current_line[item] = item_already_requested_qty_before_current_line.get(item) + ordered_qty

        # Shortage pokud inventory on hand - already requested < 0 , jinak Available.
        planned_available_po_prodeji_linky = inventory_on_hand - item_already_requested_qty_before_current_line.get(item)
        if planned_available_po_prodeji_linky < 0:
            line.append("Shortage")
        else:
            line.append("Available")
    return mp_data

def data_date_formating(data, data_headings): # Prevedeni pole datumu na date format.
    datumy = [data_headings.index(heading) for heading in data_headings if "DATE" in heading.upper()]

    for line in data: # Prevedeni pole datumu na date format.
        
        # print(datumy_indexy)
        for datum_index in datumy:      
            den, mesic, rok = line[datum_index].split("/")
            # print(rok, mesic, den)
            datum = datetime.date(int(rok), int(mesic), int(den))
            line[datum_index] = datum
    return data

def generic_database(data, headings, primary_key): # Vytvoreni dictionary databaze jednotlivych itemu s jejich linkamy a daty z reportu.
    database_dictionary = dict() # Vsechny vrcholy z reportu a jejich linky priprava dict.
    
    for heading in headings:
        if heading.upper() == str(primary_key).upper():
            pk_index = headings.index(heading) 

    for line in data:  
        key = line[pk_index]
        compile_line_dict = {} # Jednotlive linky vrcholu z dat
        data_dict = {} # Samotna data na linkach.
        # Sestaveni dat linky pro databazi.
        for i, data_field in enumerate(line):
            data_dict[headings[i]] = data_field # Sestaveni dat z linky jako dict s jmeny sloupcu zahlavi reportu jako klic.
        # print(f'Item: {item}')
        if key not in database_dictionary:# Pokud item jeste neni v databazi → vytvori ho tam s touto prvni linkou.
            compile_line_dict[1] = data_dict # Sestaveni jednotlivych linek jako dict cislo linky vrcholu : data v lince vyse.
            # print(compile_line_dict)             
            # print(f'Neni tam zatim.')
            database_dictionary[key] = compile_line_dict # Sestaveni linek predchoziho vrcholu jako dict vrchol : jeho linky s daty viz. vyse.
            # print(war_loc_database_dictionary)
        else:# Pokud item uz je v databazi..
            # print(f'Uz tam je.')
            # Pokud pridavana lokace tam jeste neni → prodat samostatnou linku.
            database_dictionary[key][len(database_dictionary[key])+1] = data_dict
            # print(war_loc_database_dictionary)
    return database_dictionary

def order_plan_database_pzn100(data, headings): # Vytvoreni dictionary PZN100 jednotlivych itemu s jejich linkamy a daty z reportu.
    database_pzn100_order_plan_dictionary = dict() # Vsechny vrcholy z PZN100 z reportu : jejich linky.

    for heading in headings:
        if heading.upper() == "ITEM":
            item_index = headings.index(heading)
        if heading.upper() == "CLUSTER" or heading.upper() == "WAREHOUSE":
            sklad_index = headings.index(heading) 

    for line in data:  
        if line[sklad_index].upper() == "PZN100":
            item = line[item_index]
            compile_line_dict = {} # Jednotlive linky vrcholu z dat
            data_dict = {} # Samotna data na linkach.
            # Sestaveni dat linky pro databazi.
            for i, data_field in enumerate(line):
                data_dict[headings[i]] = data_field # Sestaveni dat z linky jako dict s jmeny sloupcu zahlavi reportu jako klic.
            # print(f'Item: {item}')
            if item not in database_pzn100_order_plan_dictionary:# Pokud item jeste neni v databazi → vytvori ho tam s touto prvni linkou.
                compile_line_dict[1] = data_dict # Sestaveni jednotlivych linek jako dict cislo linky vrcholu : data v lince vyse.
                # print(compile_line_dict)             
                # print(f'Neni tam zatim.')
                database_pzn100_order_plan_dictionary[item] = compile_line_dict # Sestaveni linek predchoziho vrcholu jako dict vrchol : jeho linky s daty viz. vyse.
                # print(war_loc_database_dictionary)
            else:# Pokud item uz je v databazi..
                # print(f'Uz tam je.')
                # K tomuto itemu prida na konec resenou linku jako dalsi zaznam.
                database_pzn100_order_plan_dictionary[item][len(database_pzn100_order_plan_dictionary[item])+1] = data_dict
                # print(war_loc_database_dictionary)
    return database_pzn100_order_plan_dictionary

def order_plan_database_pzn105(data, headings): # Vytvoreni dictionary PZN105 jednotlivych itemu s jejich linkamy a daty z reportu.
    database_pzn105_order_plan_dictionary = dict() # Vsechny vrcholy z PZN105 z reportu : jejich linky.
    
    for heading in headings:
        if heading.upper() == "ITEM":
            item_index = headings.index(heading)
        if heading.upper() == "CLUSTER" or heading.upper() == "WAREHOUSE":
            sklad_index = headings.index(heading) 

    for line in data:  
        if line[sklad_index].upper() == "PZ5" or line[sklad_index].upper() == "PZN105":        
            item = line[item_index]
            compile_line_dict = {} # Jednotlive linky vrcholu z dat
            data_dict = {} # Samotna data na linkach.
            # Sestaveni dat linky pro databazi.
            for i, data_field in enumerate(line):
                data_dict[headings[i]] = data_field # Sestaveni dat z linky jako dict s jmeny sloupcu zahlavi reportu jako klic.
            # print(f'Item: {item}')
            if item not in database_pzn105_order_plan_dictionary:# Pokud item jeste neni v databazi → vytvori ho tam s touto prvni linkou.
                compile_line_dict[1] = data_dict # Sestaveni jednotlivych linek jako dict cislo linky vrcholu : data v lince vyse.
                # print(compile_line_dict)             
                # print(f'Neni tam zatim.')
                database_pzn105_order_plan_dictionary[item] = compile_line_dict # Sestaveni linek predchoziho vrcholu jako dict vrchol : jeho linky s daty viz. vyse.
                # print(war_loc_database_dictionary)
            else:# Pokud item uz je v databazi..
                # print(f'Uz tam je.')
                # K tomuto itemu prida na konec resenou linku jako dalsi zaznam.
                database_pzn105_order_plan_dictionary[item][len(database_pzn105_order_plan_dictionary[item])+1] = data_dict
                # print(war_loc_database_dictionary)
    return database_pzn105_order_plan_dictionary   

def order_plan_database_pzn310(data, headings): # Vytvoreni dictionary PZN310 jednotlivych itemu s jejich linkamy a daty z reportu.
    database_pzn310_order_plan_dictionary = dict() # Vsechny vrcholy z PZN310 z reportu : jejich linky.
    
    for heading in headings:
        if heading.upper() == "ITEM":
            item_index = headings.index(heading)
        if heading.upper() == "CLUSTER" or heading.upper() == "WAREHOUSE":
            sklad_index = headings.index(heading) 

    for line in data:  
        if line[sklad_index].upper() == "PZ4" or line[sklad_index].upper() == "PZN310":        
            item = line[item_index]
            compile_line_dict = {} # Jednotlive linky vrcholu z dat
            data_dict = {} # Samotna data na linkach.
            # Sestaveni dat linky pro databazi.
            for i, data_field in enumerate(line):
                data_dict[headings[i]] = data_field # Sestaveni dat z linky jako dict s jmeny sloupcu zahlavi reportu jako klic.
            # print(f'Item: {item}')
            if item not in database_pzn310_order_plan_dictionary:# Pokud item jeste neni v databazi → vytvori ho tam s touto prvni linkou.
                compile_line_dict[1] = data_dict # Sestaveni jednotlivych linek jako dict cislo linky vrcholu : data v lince vyse.
                # print(compile_line_dict)             
                # print(f'Neni tam zatim.')
                database_pzn310_order_plan_dictionary[item] = compile_line_dict # Sestaveni linek predchoziho vrcholu jako dict vrchol : jeho linky s daty viz. vyse.
                # print(war_loc_database_dictionary)
            else:# Pokud item uz je v databazi..
                # print(f'Uz tam je.')
                # K tomuto itemu prida na konec resenou linku jako dalsi zaznam.
                database_pzn310_order_plan_dictionary[item][len(database_pzn310_order_plan_dictionary[item])+1] = data_dict
                # print(war_loc_database_dictionary)
    return database_pzn310_order_plan_dictionary

def obsolete_polozky_database(data, headings): # Vytvoreni dictionary databaze obsolete polozek z reportu.
    obsolete_database_dictionary = dict() # Vsechny vrcholy z reportu a jejich linky priprava dict.
    
    for heading in headings:
        if heading.upper() == "ITEM":
            item_index = headings.index(heading)
        if heading.upper() == "CLUSTER" or heading.upper() == "WAREHOUSE":
            sklad_index = headings.index(heading)

    for line in data:  
        item = line[item_index]
        warehouse_dict = {}
        compile_line_dict = {} # Jednotlive linky vrcholu z dat
        data_dict = {} # Samotna data na linkach.
        # Sestaveni dat linky pro databazi.
        for i, data_field in enumerate(line):
            data_dict[headings[i]] = data_field # Sestaveni dat z linky jako dict s jmeny sloupcu zahlavi reportu jako klic.
        # print(f'Item: {item}')
        if item not in obsolete_database_dictionary:# Pokud item jeste neni v databazi → vytvori ho tam s timto skladem a prvni touto linkou toho skladu.
            compile_line_dict[1] = data_dict # Sestaveni dict jednotlivych linek s cislem jejich poradi a daty z dane linky ziskane vyse.
            warehouse_dict[line[sklad_index]] = compile_line_dict  # Sestaveni dict jednotlivych WHS s jejich cislem a jednotlivymi linkamy daneho skaldu ziskane vyse.          
            # print(compile_line_dict)             
            # print(f'Neni tam zatim.')
            obsolete_database_dictionary[item] = warehouse_dict # Zalozeni tohoto itemu se skaldem a daty resene linky.
            # print(war_loc_database_dictionary)
        else:# Pokud item uz je v databazi.
            # Pokud tam jeste neni sklad resene linky:
            if line[sklad_index] not in obsolete_database_dictionary[item]:
                compile_line_dict[1] = data_dict # Sestaveni dict jednotlivych linek s cislem jejich poradi a daty z dane linky ziskane vyse.
                obsolete_database_dictionary[item][line[sklad_index]] = compile_line_dict # Pridani skladu resene linky do dict reseneho itemu s resenou linku jako prvni linkou daneho skladu.
            # Pokud uz tam je sklad resene linky:
            else:
                sklad = line[sklad_index]
                linek_itemu_uz_na_danem_skladu = len(obsolete_database_dictionary[item][sklad])
                obsolete_database_dictionary[item][sklad][linek_itemu_uz_na_danem_skladu+1] = data_dict # Pridani resene linky k ostatnim linkam daneho skaldu itemu na konec.
            # print(war_loc_database_dictionary)
    return obsolete_database_dictionary

def warehouse_locations_database(data, headings): # Vytvoreni dictionary warehouse jednotlivych itemu s jejich linkamy a daty z reportu.
    war_loc_database_dictionary = dict() # Vsechny vrcholy z reportu a jejich linky priprava dict.
    # print(headings)
    for heading in headings:
        if heading.upper() == "ITEM":
            item_index = headings.index(heading)
        if heading.upper() == "ITEM LINE IN LN":
            item_line_index = headings.index(heading)
        if heading.upper() == "LOCATION":
            location_index = headings.index(heading)  
        if heading.upper() == "INVENTORY ON HAND":
            ioh_index = headings.index(heading)            

    for line in data:  
        # print(f'Line: {line}')
        item = line[item_index]
        compile_line_dict = {} # Jednotlive linky vrcholu z dat
        data_dict = {} # Samotna data na linkach.
        # Sestaveni dat linky pro databazi.
        for i, data_field in enumerate(line):
            data_dict[headings[i]] = data_field # Sestaveni dat z linky jako dict s jmeny sloupcu zahlavi reportu jako klic.
        compile_line_dict[line[item_line_index]] = data_dict # Sestaveni jednotlivych linek jako dict cislo linky vrcholu : data v lince vyse.
        # print(compile_line_dict)  

        # print(f'Item: {item}')
        if item not in war_loc_database_dictionary:# Pokud item jeste neni v databazi → vytvori ho tam s touto prvni linkou.
            # print(f'Neni tam zatim.')
            war_loc_database_dictionary[item] = compile_line_dict # Sestaveni linek predchoziho vrcholu jako dict vrchol : jeho linky s daty viz. vyse.
            # print(war_loc_database_dictionary)
        else:# Pokud item uz je v databazi..
            # print(f'Uz tam je.')
            # kouknout na vsechny linky itemu, co uz tam jsou, jestli je nejaka linka stejna lokace jako resena linka.
            for linka in war_loc_database_dictionary.get(item):
                # print(f'Linka: {war_loc_database_dictionary.get(item).get(linka)}')
                # Pokud lokace toho itemu v pridavane lince uz je v databazi u tohoto itemu (v datech je vice linek stejneho itemu a lokace), mnozstvi v pridavane
                #  lince se pricte k lince, ktera uz tam je. (sectou se mnozstvi na stejne lokaci).

                lokace_already_in_database = war_loc_database_dictionary.get(item).get(linka).get("Location").upper()
                lokace_resene_linky = line[location_index].upper()

                if lokace_already_in_database == lokace_resene_linky:
                    # Pokud lokace toho itemu v pridavane lince uz je v databazi u tohoto itemu (v datech je vice linek stejneho itemu a lokace), mnozstvi v pridavane lince se pricte k lince, ktera uz tam je. (sectou se mnozstvi na stejne lokaci).
                    war_loc_database_dictionary[item][linka][headings[ioh_index]] = str(float(war_loc_database_dictionary.get(item).get(linka).get(headings[ioh_index]).replace(",","")) + float(line[ioh_index].replace(",","")))
                    break
            else:# Pokud pridavana lokace tam jeste neni → prodat samostatnou linku.
                war_loc_database_dictionary[item][len(war_loc_database_dictionary[item])+1] = data_dict
            # print(war_loc_database_dictionary)
    return war_loc_database_dictionary

def zahlavi_vystupu_cq(zahlavi_master_planu): # Pripravi zahlavi sloupcu vystupu programu. 
    zahlavi_vystupu = list()

    for nazev in zahlavi_master_planu: # Nalezeni indexu jednotlivych nazvu sloupcu ze zahlavi Master planu.
        if nazev.upper() == "SALES ORDER": # Najit sloupec Sales order.
            sales_order_sloupec_index = zahlavi_master_planu.index(nazev)
        elif nazev.upper() == "PROJECT": # Najit Project.
            project_sloupec_index = zahlavi_master_planu.index(nazev)
        elif nazev.upper() == "ITEM": # Najit sloupec Item.
            item_sloupec_index = zahlavi_master_planu.index(nazev)
        elif nazev.upper() == "DESCRIPTION": # Najit sloupec Description.
            description_sloupec_index = zahlavi_master_planu.index(nazev)        
        elif nazev.upper() == "PLANNED DELIVERY DATE": # Najit sloupec PDD.
            pdd_sloupec_index = zahlavi_master_planu.index(nazev)        
        elif nazev.upper() == "CUSTOMER REQ DATE": # Najit sloupec CRD.
            crd_sloupec_index = zahlavi_master_planu.index(nazev)        
        elif nazev.upper() == "SUPPLIER PROMISED DATE": # Najit sloupec SPD.
            spd_sloupec_index = zahlavi_master_planu.index(nazev)        
        elif nazev.upper() == "ORDERED QTY": # Najit sloupec Ordered Qty.
            ordered_qty_sloupec_index = zahlavi_master_planu.index(nazev)
        elif nazev.upper() == "ROW IN MASTER PLAN": # Najit sloupec Row in Master plan.
            mp_row_loupec_index = zahlavi_master_planu.index(nazev)                  

    # Pripraveni zahlavi vystupu programu v danem poradi.
    zahlavi_vystupu.append(zahlavi_master_planu[item_sloupec_index])
    zahlavi_vystupu.append(zahlavi_master_planu[mp_row_loupec_index])
    zahlavi_vystupu.append(zahlavi_master_planu[description_sloupec_index])
    zahlavi_vystupu.append(zahlavi_master_planu[sales_order_sloupec_index])
    zahlavi_vystupu.append(zahlavi_master_planu[pdd_sloupec_index])
    zahlavi_vystupu.append(zahlavi_master_planu[crd_sloupec_index])
    zahlavi_vystupu.append(zahlavi_master_planu[spd_sloupec_index])
    zahlavi_vystupu.append("Buyer")
    zahlavi_vystupu.append("Supply LT[WD]")    
    zahlavi_vystupu.append(zahlavi_master_planu[ordered_qty_sloupec_index])
    zahlavi_vystupu.append("Sum req. Qty Item+PDD")
    zahlavi_vystupu.append("Planned available Qty on PZN105 at PDD")
    zahlavi_vystupu.append("PZN105 purchase orders, ktere musi dnes prijit, aby PA odpovidalo")
    zahlavi_vystupu.append("Planned available Qty on PZN100 at PDD")
    zahlavi_vystupu.append("PZN100 purchase orders, ktere musi dnes prijit, aby PA odpovidalo")
    zahlavi_vystupu.append("Already requested in Tabulka prevodu na PZN105")
    zahlavi_vystupu.append("PZN 105 Nejblizsi datum + mnozstvi: kdy PA >= 0")
    zahlavi_vystupu.append("PZN 100 Nejblizsi datum + mnozstvi: kdy PA >= 0")
    zahlavi_vystupu.append("Mozno prevest z PZN100 aniz by se ohrozily budouci linky na PZN100?")
    zahlavi_vystupu.append("Prevest z lokace")
    return zahlavi_vystupu

def shortage_linky_master_planu_cq(mp_data, zahlavi_master_planu, do_datumu_proverit_master_plan, order_plan_pzn_100, order_plan_pzn_105): # Projde vsechny linky Master planu a vybere na zaklade kriterii shortage linky k provereni. Vysledek vrati jako seznam linek Master planu serazeny podle linky Master planu a itemu.
    
    shortage_linky = list()

    for nazev in zahlavi_master_planu: # Najit indexy nazvu sloupcu z Master planu.
        if nazev.upper() == "AVAILABILITY": # Najit sloupec Availability.
            availability_sloupec_index = zahlavi_master_planu.index(nazev)
        elif nazev.upper() == "SALES ORDER": # Najit sloupec Sales order.
            sales_order_sloupec_index = zahlavi_master_planu.index(nazev)
        elif nazev.upper() == "PROJECT": # Najit Project.
            project_sloupec_index = zahlavi_master_planu.index(nazev)
        elif nazev.upper() == "ITEM": # Najit sloupec Item.
            item_sloupec_index = zahlavi_master_planu.index(nazev)
        elif nazev.upper() == "DESCRIPTION": # Najit sloupec Description.
            description_sloupec_index = zahlavi_master_planu.index(nazev)        
        elif nazev.upper() == "PLANNED DELIVERY DATE": # Najit sloupec PDD.
            pdd_sloupec_index = zahlavi_master_planu.index(nazev)        
        elif nazev.upper() == "CUSTOMER REQ DATE": # Najit sloupec CRD.
            crd_sloupec_index = zahlavi_master_planu.index(nazev)        
        elif nazev.upper() == "SUPPLIER PROMISED DATE": # Najit sloupec SPD.
            spd_sloupec_index = zahlavi_master_planu.index(nazev)        
        elif nazev.upper() == "ORDERED QTY": # Najit sloupec Ordered Qty.
            ordered_qty_sloupec_index = zahlavi_master_planu.index(nazev)
        elif nazev.upper() == "ROW IN MASTER PLAN": # Najit sloupec s poradim linky v Master planu.
            line_row_sloupec_index = zahlavi_master_planu.index(nazev)


    datumy_sloupce_indexy = [pdd_sloupec_index, crd_sloupec_index, spd_sloupec_index]

    for line in mp_data: # Sestaveni.  
        if line[availability_sloupec_index] != None and line[availability_sloupec_index].upper() == "SHORTAGE": # Overeni ze je linka Shortage.
            if line[project_sloupec_index] == "0": # Kontrola anonymni polozky.
                for datum_linky in [line[index_datumu] for index_datumu in datumy_sloupce_indexy]: # Kontrola datumu. Proveruje PDD, CRD i SPD linky.                    
                    if not neplatne_datum_ln(datum_linky) and datum_linky <= do_datumu_proverit_master_plan and datum_linky >= dnesni_datum():

                        # Sestaveni linky z vybranych bunek. Item, linka Master Planu, Description, Sales order, PDD, CRD, SPD, Nakupci, Ordered Qty linky,   
                        list_linky = list() 
                        
                        list_linky.append(line[item_sloupec_index])
                        list_linky.append(line[line_row_sloupec_index])
                        list_linky.append(line[description_sloupec_index])
                        list_linky.append(line[sales_order_sloupec_index])
                        list_linky.append(line[pdd_sloupec_index].strftime("%d/%m/%Y").replace("/", "."))
                        list_linky.append(line[crd_sloupec_index].strftime("%d/%m/%Y").replace("/", "."))
                        list_linky.append(line[spd_sloupec_index].strftime("%d/%m/%Y").replace("/", "."))
                        # Pridani sloupce buyer
                        # Pokud je alespon v jednom order planu vrchol, vezme udaj nakupci z neho, pokud ne, da hodnotu 0.
                        if order_plan_pzn_100.get(line[item_sloupec_index].strip()) or order_plan_pzn_105.get(line[item_sloupec_index].strip()):
                            buyer = order_plan_pzn_100.get(line[item_sloupec_index].strip()).get(1).get("Buyer") if order_plan_pzn_100.get(line[item_sloupec_index].strip()) else order_plan_pzn_105.get(line[item_sloupec_index].strip()).get(1).get("Buyer")
                        else:
                            buyer = "N/A"
                        list_linky.append(buyer)                        
                        # Pridani sloupce Supply LT
                        if order_plan_pzn_100.get(line[item_sloupec_index].strip()) or order_plan_pzn_105.get(line[item_sloupec_index].strip()):
                            supply_lt = order_plan_pzn_100.get(line[item_sloupec_index].strip()).get(1).get("Supply LT [work days]") if order_plan_pzn_100.get(line[item_sloupec_index].strip()) else order_plan_pzn_105.get(line[item_sloupec_index].strip()).get(1).get("Supply LT [work days]")
                        else:
                            supply_lt = "N/A"                        
                        list_linky.append(supply_lt)
                        
                        list_linky.append(line[ordered_qty_sloupec_index])

                        # Pridani linky k ostatnim shortage linkam.    
                        shortage_linky.append(list_linky)         
                        break # Ukoncuje for smycku kontrolujici platne datumy. (Staci najit jeden platny datum v lince, abu se pridala. Pal break aby se neopakovala stejna linka.
    shortage_linky.sort() # Seradi linky podle itemu a linky Master planu.
    return shortage_linky

def seznam_itemu_pro_order_plany_cq(mp_data, zahlavi_master_planu, do_datumu_proverit_master_plan): # Vynda z dat master planu seznam itemu k provereni order planu a odstrani duplicity.
    
    seznam_itemu_pro_order_plany = set()

    for nazev in zahlavi_master_planu: # Najit indexy nazvu sloupcu z Master planu.
        if nazev.upper() == "AVAILABILITY": # Najit sloupec Availability.
            availability_sloupec_index = zahlavi_master_planu.index(nazev)
        elif nazev.upper() == "PROJECT": # Najit Project.
            project_sloupec_index = zahlavi_master_planu.index(nazev)
        elif nazev.upper() == "ITEM": # Najit sloupec Item.
            item_sloupec_index = zahlavi_master_planu.index(nazev)       
        elif nazev.upper() == "PLANNED DELIVERY DATE": # Najit sloupec PDD.
            pdd_sloupec_index = zahlavi_master_planu.index(nazev)        
        elif nazev.upper() == "CUSTOMER REQ DATE": # Najit sloupec CRD.
            crd_sloupec_index = zahlavi_master_planu.index(nazev)        
        elif nazev.upper() == "SUPPLIER PROMISED DATE": # Najit sloupec SPD.
            spd_sloupec_index = zahlavi_master_planu.index(nazev)

    datumy_sloupce_indexy = [pdd_sloupec_index, crd_sloupec_index, spd_sloupec_index]

    for line in mp_data: # Sestaveni seznamu.  
        if line[availability_sloupec_index] != None and line[availability_sloupec_index].upper() == "SHORTAGE": # Overeni ze je linka Shortage.
            if line[project_sloupec_index] == "0": # Kontrola anonymni polozky.
                for datum_linky in [line[index_datumu] for index_datumu in datumy_sloupce_indexy]: # Kontrola datumu, zda je v rozmezi proverovanych datumu zadaneho uzivatelem. Proveruje PDD, CRD i SPD linky.                    
                    if not neplatne_datum_ln(datum_linky) and datum_linky <= do_datumu_proverit_master_plan and datum_linky >= dnesni_datum():  
                        
                        # Pokud linka splni podminky vyse → item z ni je pridan do setu itemu k provereni.
                        seznam_itemu_pro_order_plany.add(line[item_sloupec_index])    
    return seznam_itemu_pro_order_plany
