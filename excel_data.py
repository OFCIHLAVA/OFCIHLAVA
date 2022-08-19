import datetime


# Master plan
def zahlavi_master_planu(excel_sheet): # Ziskani nazvu sloupcu jako list. Sales order, Item, Description, PDD, CRD, SPD, Ordered Qty, Availability    
    zahlavi_master_planu = []
    for c in range(1,min(excel_sheet.max_column+1, 16384)):
        column_name = excel_sheet.cell(1, c).value
        if column_name not in zahlavi_master_planu and column_name:
            zahlavi_master_planu.append(column_name)
    return zahlavi_master_planu

def zahlavi_vystupu_excel(excel_sheet, zahlavi_master_planu): # Pripravi zahlavi sloupcu vystupu programu.
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

    # Pripraveni zahlavi vystupu programu v danem poradi.
    zahlavi_vystupu.append(excel_sheet.cell(1, item_sloupec_index+1).value)
    zahlavi_vystupu.append("Master plan row")
    zahlavi_vystupu.append(excel_sheet.cell(1, description_sloupec_index+1).value)
    zahlavi_vystupu.append(excel_sheet.cell(1, sales_order_sloupec_index+1).value)
    zahlavi_vystupu.append(excel_sheet.cell(1, pdd_sloupec_index+1).value)
    zahlavi_vystupu.append(excel_sheet.cell(1, crd_sloupec_index+1).value)
    zahlavi_vystupu.append(excel_sheet.cell(1, spd_sloupec_index+1).value)
    zahlavi_vystupu.append("Buyer")
    zahlavi_vystupu.append("Supply LT[WD]")    
    zahlavi_vystupu.append(excel_sheet.cell(1, ordered_qty_sloupec_index+1).value)
    zahlavi_vystupu.append("Sum req. Qty Item+PDD")
    zahlavi_vystupu.append("Planned available Qty on PZN105 at PDD")
    zahlavi_vystupu.append("Planned available Qty on PZN100 at PDD")
    zahlavi_vystupu.append("Already requested in Tabulka prevodu na PZN105")
    zahlavi_vystupu.append("PZN 105 Nejblizsi datum + mnozstvi: kdy PA >= 0.")
    zahlavi_vystupu.append("PZN 100 Nejblizsi datum + mnozstvi: kdy PA >= 0.")
    zahlavi_vystupu.append("Mozno prevest z PZN100 aniz by se ohrozily budouci linky na PZN100?")
    return zahlavi_vystupu

def dnesni_datum(): # Vrati dnesni datum ve formatu rrrr-mm-dd.
    dnesni_datum = datetime.date.today()
    return dnesni_datum

def neplatne_datum_ln(datum): # Zkontroluje, zda datum neni "pocatecni datum" v LN.
    if datum == datetime.date(1970, 1, 1):
        return True
    else:
        return False

def do_datumu_proverit_master_plan(kal_dnu_k_provereni_ode_dneska): # Do jakeho datumu se maji proverovat shortage linky v Master planu ve formatu rrrr-mm-dd.
    posun_dnu = datetime.timedelta(kal_dnu_k_provereni_ode_dneska)
    do_datumu_proverit_master_plan = dnesni_datum() + posun_dnu
    return do_datumu_proverit_master_plan

def seznam_itemu_pro_order_plany_excel(excel_sheet, zahlavi_master_planu, do_datumu_proverit_master_plan): # Vynda z dat master planu seznam itemu k provereni order planu a odstrani duplicity.
    
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
    
    # set itemu z linek Master planu k provereni order planu.
    for row in range(2 ,min(1048576, excel_sheet.max_row+1)): # Sestaveni seznamu.  
        if excel_sheet.cell(row, availability_sloupec_index+1).value != None and excel_sheet.cell(row, availability_sloupec_index+1).value.upper() == "SHORTAGE": # Overeni ze je linka Shortage.
            if excel_sheet.cell(row, project_sloupec_index+1).value == "": # Kontrola anonymni polozky.
                for datum_linky in [excel_sheet.cell(row, c+1).value.date() for c in datumy_sloupce_indexy]: # Kontrola datumu, zda je v rozmezi proverovanych datumu zadaneho uzivatelem. Proveruje PDD, CRD i SPD linky.
                    if not neplatne_datum_ln(datum_linky) and datum_linky <= do_datumu_proverit_master_plan and datum_linky >= dnesni_datum():
                        
                        # Pokud linka splni podminky vyse → item z ni je pridan do setu itemu k provereni.
                        seznam_itemu_pro_order_plany.add(excel_sheet.cell(row, item_sloupec_index+1).value)
    return seznam_itemu_pro_order_plany            

def shortage_linky_master_planu_excel(excel_sheet, zahlavi_master_planu, do_datumu_proverit_master_plan, order_plan_pzn_100, order_plan_pzn_105): # Projde vsechny linky Master planu a vybere na zaklade kriterii shortage linky k provereni. Vysledek vrati jako seznam linek Master planu serazeny podle linky Master planu a itemu.
    
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

    datumy_sloupce_indexy = [pdd_sloupec_index, crd_sloupec_index, spd_sloupec_index]


    for row in range(2 ,min(1048576, excel_sheet.max_row+1)): # Sestaveni.  
        if excel_sheet.cell(row, availability_sloupec_index+1).value != None and excel_sheet.cell(row, availability_sloupec_index+1).value.upper() == "SHORTAGE": # Overeni ze je linka Shortage.
            if excel_sheet.cell(row, project_sloupec_index+1).value == "": # Kontrola anonymni polozky.
                for datum_linky in [excel_sheet.cell(row, c+1).value.date() for c in datumy_sloupce_indexy]: # Kontrola datumu. Proveruje PDD, CRD i SPD linky.
                    if not neplatne_datum_ln(datum_linky) and datum_linky <= do_datumu_proverit_master_plan and datum_linky >= dnesni_datum():

                        # Sestaveni linky z vybranych bunek. Item, linka Master Planu, Description, Sales order, PDD, CRD, SPD, Nakupci, Ordered Qty linky,   
                        list_linky = list() 
                        
                        list_linky.append(excel_sheet.cell(row, item_sloupec_index+1).value)
                        list_linky.append(row)
                        list_linky.append(excel_sheet.cell(row, description_sloupec_index+1).value)
                        list_linky.append(excel_sheet.cell(row, sales_order_sloupec_index+1).value)
                        list_linky.append(excel_sheet.cell(row, pdd_sloupec_index+1).value.strftime("%d/%m/%Y").replace("/", "."))
                        list_linky.append(excel_sheet.cell(row, crd_sloupec_index+1).value.strftime("%d/%m/%Y").replace("/", "."))
                        list_linky.append(excel_sheet.cell(row, spd_sloupec_index+1).value.strftime("%d/%m/%Y").replace("/", "."))
                        # Pridani sloupce buyer
                        # Pokud je alespon v jednom order planu vrchol, vezme udaj nakupci z neho, pokud ne, da hodnotu 0.
                        if order_plan_pzn_100.get(excel_sheet.cell(row, item_sloupec_index+1).value.strip()) or order_plan_pzn_105.get(excel_sheet.cell(row, item_sloupec_index+1).value.strip()):
                            buyer = order_plan_pzn_100.get(excel_sheet.cell(row, item_sloupec_index+1).value.strip()).get(1).get("Buyer") if order_plan_pzn_100.get(excel_sheet.cell(row, item_sloupec_index+1).value.strip()) else order_plan_pzn_105.get(excel_sheet.cell(row, item_sloupec_index+1).value.strip()).get(1).get("Buyer")
                        else:
                            buyer = "N/A"
                        list_linky.append(buyer)                        
                        # Pridani sloupce Supply LT
                        if order_plan_pzn_100.get(excel_sheet.cell(row, item_sloupec_index+1).value.strip()) or order_plan_pzn_105.get(excel_sheet.cell(row, item_sloupec_index+1).value.strip()):
                            supply_lt = order_plan_pzn_100.get(excel_sheet.cell(row, item_sloupec_index+1).value.strip()).get(1).get("Supply LT [work days]") if order_plan_pzn_100.get(excel_sheet.cell(row, item_sloupec_index+1).value.strip()) else order_plan_pzn_105.get(excel_sheet.cell(row, item_sloupec_index+1).value.strip()).get(1).get("Supply LT [work days]")
                        else:
                            supply_lt = "N/A"                        
                        list_linky.append(supply_lt)
                        
                        list_linky.append(excel_sheet.cell(row, ordered_qty_sloupec_index+1).value)

                        # Pridani linky k ostatnim shortage linkam.    
                        shortage_linky.append(list_linky)         
                        break # Ukoncuje for smycku kontrolujici platne datumy. (Staci najit jeden platny datum v lince, abu se pridala. Pal break aby se neopakovala stejna linka.
    shortage_linky.sort() # Seradi linky podle itemu a linky Master planu.
    return shortage_linky

def doplneni_sum_ordered_qty_do_vystupu(shortage_linky, zahlavi_vystupu): # Pro kazdou linku vytazenou z Master planu prida na konec linky sumu pozadavku Qty v dany den pro dany vrchol a PDD linky (suma vsech stejnych vrcholu v dane PDD)
    for line in shortage_linky:
        vrchol = line[zahlavi_vystupu.index("Item")]
        pdd = line[zahlavi_vystupu.index("Planned Delivery Date")]
        requested_qty_vrchol_pdd = 0
        
        if vrchol == "Item":
            continue        
        for linka in shortage_linky: # Projde vsechny linky a nacte pozadavky na Qty pro dane PDD datum pro dany vrchol.
            if linka[zahlavi_vystupu.index("Item")] == vrchol and linka[zahlavi_vystupu.index("Planned Delivery Date")] == pdd:
                requested_qty_vrchol_pdd += float(linka[zahlavi_vystupu.index("Ordered Qty")])
        line.append(float(requested_qty_vrchol_pdd)) # Pripoji udaj na konec linky.

def doplneni_zahlavi_do_vystupu(shortage_linky, zahlavi_vystupu):
    shortage_linky.insert(0, zahlavi_vystupu)

def doplneni_already_zadano_do_vystupu(excel_sheet, shortage_linky, zahlavi_vystupu, obsolete_days): # Pro kazdou linku vytazenou z Master planu prida na konec linky sumu Qty, o kterou je zazadano v tabulce prevodu ale jeste nebylo zpracovane.
    
    neplatne_pozadavky_po_kolika_dnech = datetime.timedelta(obsolete_days) # Po kolika dnech uz nezpracovane pozadavky v tabulce prevodu povazujeme za naplatne.    
    
    # Cislo sloupcu v zahlavi tabulky prevodu.
    item_col = 2
    qty_col = 3
    from_whs_col = 7
    to_whs_col = 8
    date_created_col = 10
    who_processed_col = 11

    for line in shortage_linky:
        vrchol = str(line[zahlavi_vystupu.index("Item")])
        if vrchol == "Item":
            continue
        vrchol_uz_zadano_k_prevodu = 0
        for row in range(7 ,min(1048576, excel_sheet.max_row+1)):
            cell_item = excel_sheet.cell(row, item_col)
            cell_who_processed = excel_sheet.cell(row, who_processed_col)
            cell_date_created = excel_sheet.cell(row, date_created_col)
            cell_qty = excel_sheet.cell(row, qty_col)
            cell_from_whs = excel_sheet.cell(row, from_whs_col)
            cell_to_whs = excel_sheet.cell(row, to_whs_col)
            if str(cell_who_processed.value).strip() == "None": # Kontrola, ze linka zatim nebyla prevedena.                          
                # print(row, str(cell_who_processed.value).strip(), type(str(cell_who_processed.value).strip()))
                # Kontrola stari pozadavku.
                if str(type(cell_date_created.value)) == "<class 'str'>": 
                    try: # Pokud datum neni ve formatu datumu, pokusim se ho rozdelit podle "/" resp "." a z vysledku vzit rok, mesic a datum a udelat z toho date promenou.
                        if "/" in cell_date_created.value[0:4]:
                            splitted_date = cell_date_created.value.split("/")
                            splitted_date = datetime.date(int(splitted_date[2][0:4]),int(splitted_date[1]),int(splitted_date[0]))
                        if "." in cell_date_created.value[0:4]:
                            splitted_date = cell_date_created.value.split(".")
                            datum_created = datetime.date(int(splitted_date[2][0:4]),int(splitted_date[1]),int(splitted_date[0]))
                    except:
                        print(f'POZOR! Datum Kdy zadano v souboru Prevody na PZN105 na radce {row} je v nespravnem formatu. Nebylo mozno overit stav na teto radce.')
                        continue                  
                elif str(type(cell_date_created.value)) == "<class 'datetime.datetime'>":
                    datum_created = cell_date_created.value.date()

                if datum_created > dnesni_datum() - neplatne_pozadavky_po_kolika_dnech:
                    if (str(cell_from_whs.value).strip()).upper() == "PZN100" and (str(cell_to_whs.value).strip()).upper() == "PZN105": # Kontrola, ze linka se tyka prevodu z PZN100 na PZN105.  
                        if str(cell_item.value).strip() == vrchol:
                            vrchol_uz_zadano_k_prevodu += int(cell_qty.value)
        line.append(vrchol_uz_zadano_k_prevodu)