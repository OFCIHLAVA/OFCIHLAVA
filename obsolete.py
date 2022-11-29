from funkce.cq_data import data_date_formating, data_import, data_headings, generic_database,import_data_cleaning, order_plan_database_pzn100, order_plan_database_pzn105, order_plan_database_pzn310, obsolete_polozky_database
from funkce.funkce_prace import dotaz_pn_program, nacteni_databaze_boud_pro_dotaz, programy_boud
from funkce.prevody_dotazy import planned_available_na_skladu, po_in_process_skladu, realny_demand_na_skladu

import datetime
import sys
import time
import openpyxl as excel


# Priprava databazi pro dotazovani v jakych boudach je item obsazen.
path_kusovniky_databaze = 'Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\5_SFExBFExMIX\\databaze\\databaze boud s kusovniky.txt'
path_program_databaze = 'Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\5_SFExBFExMIX\\databaze\\seznam programu.txt'

databaze_kusovniku_pro_dotaz = nacteni_databaze_boud_pro_dotaz(path_kusovniky_databaze)
seznam_boud_z_databaze_kusovniku = [key for key in databaze_kusovniku_pro_dotaz]
print("Databaze bud s kusovniky nactena a pripravena pro dotazovani . . .")

kvp_programy_pro_dotaz = programy_boud(path_program_databaze)
seznam_boud_z_databaze_programu = [key for key in kvp_programy_pro_dotaz]
print("Databaze programu vsech boud nacetenaa pripravena pro dotazovani . . .\n")

obsolete_file_import = "Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\2_Obsolete analýza\\polozky_proverit\\obsolete polozky.txt"
obsolete_order_plan_import = "Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\2_Obsolete analýza\\order_plany\\order_plan_100_105_310.txt"
obsolete_last_transactions_import = "Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\2_Obsolete analýza\\posledni_transakce\\obselete_last_transaction.txt"

# Nacteni poctu boud z txt souboru vytazeneho z PSR a doplneni informace, jestli jsou nektere linky itemy retired.
# 
psr_status_file_import = "Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\2_Obsolete analýza\\psr in service status\\boudy_service_status.txt"
psr_status_data = data_import(psr_status_file_import)
psr_status_data_headings = data_headings(psr_status_data)
psr_status_data = import_data_cleaning(psr_status_data)

psr_status_databaze = generic_database(psr_status_data, psr_status_data_headings)

# for bouda, linky in psr_status_databaze.items():
#     print(bouda)
#     for cislo, linka in linky.items():
#         print(cislo, linka,sep=": ")      

### OBSOLETE POLOZKY

input(f'\nKROK 1:\n\nSpust report:\nusers/Ondrej.Rott/Obsolete analyza/Oboslete analyza POLOZKY.eq\na vysledek reportu uloz d slozky:\nY:/Departments/Sales and Marketing/Aftersales/11_PLANNING/23_Python_utilities/Obsolete analýza/polozky_proverit pod nazvem:\nobsolete polozky.txt\nPote pokracuj stiskem ENTER . . . ')

# Import dat obsolete polozek.
try:
    obsolete_polozky_data = data_import(obsolete_file_import)
except FileNotFoundError:
    print(f'\nPOZOR! Ve slozce: Y:/Departments/Sales and Marketing/Aftersales/11_PLANNING/23_Python_utilities/Obsolete analýza/polozky_proverit NEBYL nalezen soubor: obsolete polozky.txt\nZkontroluj, zda je tam soubor ulozen a zkus znovu.')
    input(f'Ukonci program klavesou ENTER. . .')
    sys.exit()

obsolete_polozky_data_headings = data_headings(obsolete_polozky_data)
obsolete_polozky_data = import_data_cleaning(obsolete_polozky_data)

# Vytvoreni databaze obsolete polozek.
obsolete_polozky_databaze = obsolete_polozky_database(obsolete_polozky_data, obsolete_polozky_data_headings)

# Vytovreni seznamu obsolete polozek do cq reportu order planu.
obs_polozky_seznam = [item for item in obsolete_polozky_databaze]

obs_polozky_seznam = set(obs_polozky_seznam)
# print(f'Seznam OBSOLETE polozek k provereni:\n')
# for item in obs_polozky_seznam:
#     print(item)
with open("seznam obsolete polozek proverit.txt", "w", encoding='UTF-8') as op:
    for item in obs_polozky_seznam:
        op.write(f'{item}\n')
        
print(f'\nCelkem polozek k provereni: {len(obs_polozky_seznam)}.\nSeznam polozek do dalsich CQ reportu je ulozen ve slozce programu pod nazvem seznam obsolete polozek proverit.txt.')    

### LAST TRANSACTIONS
input(f'\nKROK 2:\n\nVytvoril se txt soubor seznam obsolete polozek proverit.txt. Vloz seznam polozek z tohoto souboru do reportu:\nusers/Ondrej.Rott/Obsolete analyza/Issued for spares.eq\na vysledek reportu uloz d slozky:/nY:/Departments/Sales and Marketing/Aftersales/11_PLANNING/23_Python_utilities/Obsolete analýza/posledni_transakce pod nazvem:\nobselete_last_transaction.txt\nPote pokracuj stiskem ENTER . . . ')

# Import dat last transactions obsolete polozek.
try:
    obsolete_last_transactions_data = data_import(obsolete_last_transactions_import)
except FileNotFoundError:
    print(f'\nPOZOR! Ve slozce: Y:/Departments/Sales and Marketing/Aftersales/11_PLANNING/23_Python_utilities/Obsolete analýza/posledni_transakce NEBYL nalezen soubor: obselete_last_transaction.txt\nZkontroluj, zda je tam soubor ulozen a zkus znovu.')
    input(f'Ukonci program klavesou ENTER. . .')
    sys.exit()

obsolete_last_transactions_data_headings = data_headings(obsolete_last_transactions_data)
obsolete_last_transactions_data = import_data_cleaning(obsolete_last_transactions_data)
obsolete_last_transactions_data = data_date_formating(obsolete_last_transactions_data, obsolete_last_transactions_data_headings)

obsolete_last_transactions_databaze = obsolete_polozky_database(obsolete_last_transactions_data, obsolete_last_transactions_data_headings)


#v for item, sklady in obsolete_last_transactions_databaze.items():
#v     if item == "243545-5":
#v         print(f'{item}')
#v         for sklad, linky in sklady.items():
#v             print(sklad)
#v             for linka, data in linky.items():
#v                 print(f'{linka} :  {data}')


### ORDER PLANY

input(f'\nKROK 3:\n\nVytvoril se txt soubor seznam obsolete polozek proverit.txt. Vloz seznam polozek z tohoto souboru do reportu:\nusers/Ondrej.Rott/Obsolete analyza/PZN100+105+310_order_plan.eq\na vysledek reportu uloz d slozky:/nY:/Departments/Sales and Marketing/Aftersales/11_PLANNING/23_Python_utilities/Obsolete analýza/order_plany pod nazvem:\norder_plan_100_105_310.txt\nPote pokracuj stiskem ENTER . . . ')

# Import dat order planu obsolete polozek.
try:
    obsolete_polozky_order_plan_data = data_import(obsolete_order_plan_import)
except FileNotFoundError:
    print(f'\nPOZOR! Ve slozce: Y:/Departments/Sales and Marketing/Aftersales/11_PLANNING/23_Python_utilities/Obsolete analýza/order_plany NEBYL nalezen soubor: order_plan_100_105_310.txt\nZkontroluj, zda je tam soubor ulozen a zkus znovu.')
    input(f'Ukonci program klavesou ENTER. . .')
    sys.exit()

obsolete_polozky_order_plan_data_headings = data_headings(obsolete_polozky_order_plan_data)
obsolete_polozky_order_plan_data = import_data_cleaning(obsolete_polozky_order_plan_data)
obsolete_polozky_order_plan_data = data_date_formating(obsolete_polozky_order_plan_data, obsolete_polozky_order_plan_data_headings)

# Vytvoreni databaze order planu obsolete polozek.
order_plan_100 = order_plan_database_pzn100(obsolete_polozky_order_plan_data, obsolete_polozky_order_plan_data_headings)
order_plan_105 = order_plan_database_pzn105(obsolete_polozky_order_plan_data, obsolete_polozky_order_plan_data_headings)
order_plan_310 = order_plan_database_pzn310(obsolete_polozky_order_plan_data, obsolete_polozky_order_plan_data_headings)

### PROGRAM

# 1. Pro kazdou obsolete polozku zjistit volne QTY (bez alokace) na skladech 105 a 310.

# A) 105.
# print(obsolete_polozky_databaze)

input(f'\n\nProgram ready.\nZabere to cca {round(len(obs_polozky_seznam)*0.1549)} sekund / {round(len(obs_polozky_seznam)*0.1549/60,1)} minut\n\nSpust stiskem ENTER . . .')
start_time = time.time()

print(f'Item|IOH 105|planned available 105 (pegging)|Last transaction 105|IOH 310|planned available 310 (pegging)|Last transaction 310|Real demand 100 (pegging)|Sum qty already incoming Purchase orders na 100|Obsazeno v boudach|Pocet letajicich boud s itemem (PSR)|Obsazeno v RETIRED boudach|Pocet RETIRED boud s itemem (PSR)|Akce 105|Akce310')
with open("output.txt", "w", encoding="utf-8") as output:
    output.write(f'Item|IOH 105|planned available 105 (pegging)|Last transaction 105|IOH 310|planned available 310 (pegging)|Last transaction 310|Real demand 100 (pegging)|Sum qty already incoming Purchase orders na 100|Obsazeno v boudach|Pocet letajicich boud s itemem (PSR)|Obsazeno v RETIRED boudach|Pocet RETIRED boud s itemem (PSR)|Akce 105|Akce310\n')
# 1. Projiti databaze vsech obsolete polozek na 105 a 310:
for item, sklady in obsolete_polozky_databaze.items():    
    # print(item)
    # Reset na nove itemy

    pole_100_real_demand = "reset hodnota"
    pole_100_pur_o_in_process = "reset hodnota"
    item_stock_op_100 = 0
    pocet_stock_linek_itemu_100 = 0    
    
    ioh_105 = "reset hodnota"
    item_stock_op_105 = 0
    pocet_stock_linek_itemu_105 = 0
    pole_105_pa = "reset hodnota"
    pole_105_last_t = "reset hodnota"
    last_trans_date_105 = datetime.date(1,1,1)
    transaction_date_105 = datetime.date(1,1,1)

    ioh_310 = "reset hodnota"
    item_stock_op_310 = 0
    pocet_stock_linek_itemu_310 = 0
    pole_310_pa = "reset hodnota"
    pole_310_last_t = "reset hodnota"
    last_trans_date_310 = datetime.date(1,1,1)
    transaction_date_310 = datetime.date(1,1,1)

    v_boudach = "reset hodnota"

    v_retired_boudach = "reset hodnota"


    # Prochazeni 105 a 310 op na planned available.
    # 105 
    if "PZN105" in sklady:    
        sklad = obsolete_polozky_databaze.get(item).get("PZN105")
        # 105 IOH z I360
        # print(f'item {item} ma i360 qty na 105.')
        # projde linkty a nacete mnozstvi na skladu z nich jako ioh.
        for linka in sklad:
            ioh_105 = sklad.get(linka).get("Inventory on hand")
            ioh_105 = float(ioh_105.replace(",",""))
                # Planned available 105.
        op_105_itemu = order_plan_105.get(item)
        # Nektere itemy maji Inventory 360 data, ale nemaji order plan data (nefunkcni pegging) → tyto z kontroly vynechat.
        if op_105_itemu != None:
            # print(f'item {item} ma op data 105.')
            # Zkontrolovat STOCK ze neni divny odpovida ioh.
            for linka, data in op_105_itemu.items():
                order_type_105 = data.get("Order type txt")
                transaction_qty_105 = float(data.get("Quantity").replace(",",""))
                if order_type_105.upper() == "STOCK":
                    item_stock_op_105 = transaction_qty_105
                    pocet_stock_linek_itemu_105 += 1
            # Pokud stock linek je mene nez 2 je to OK → overit jeste, jestli si stock odpovida s IOH.
            # print(f'item {item} ma pocet STOC linek na op 105 {pocet_stock_linek_itemu_105}')
            if pocet_stock_linek_itemu_105 <2:
                # print(f'item {item} STOCK linky OK')
                # Kontrola zda si odpovida IOH z inventory dat a STOCK z op dat.
                # print(f'item {item} ma STOCK ioh z op 105 {item_stock_op_105} a ioh z I360 {ioh_105}')
                if ioh_105 == item_stock_op_105:
                    # print(f'item {item} IOH jsou stejne OK')
                    # Pokud ano, spocitat planned available na danem sklade.
                    pole_105_pa = planned_available_na_skladu(item, op_105_itemu)
                    # print(f'item {item} planned available op 105 {pole_105_pa}')
                else:
                    pole_105_pa = f'POZOR: item {item} ma jinou hodnotu STOCK transakce v pegging datech ({item_stock_op_105}) a v inventory datech ({ioh_105})!'                        
            else:                  
                pole_105_pa = f'POZOR! Pocet STOCK linek itemu {item} = {pocet_stock_linek_itemu_105}. Nebude kontrolovan.'           
        else:
            pole_105_pa = f'POZOR! {item} ma warehouse data 105, ale v order planu 105 nic nema (rozbity pegging u itemu). Nebude kontrolovan.'               
    # print(f'item {item} i360 qty na 105 = {ioh_105}.')                
    # Pokud item ma nejake QTY 0 ve whs 105 v i360, ale pri tom ma op stock vetsi nez jedna, je to diven.
    else:
        ioh_105 = 0
        op_105_itemu = order_plan_105.get(item)
        # Nektere itemy maji Inventory 360 data, ale nemaji order plan data (nefunkcni pegging) → tyto z kontroly vynechat.
        if op_105_itemu != None:
            # print(f'item {item} ma op data 105.')
            # Zkontrolovat STOCK ze neni divny odpovida ioh.
            pole_105_pa = planned_available_na_skladu(item, op_105_itemu)
        # Pokud neni nic na skladu PZN 105 I360 LN ani nejsou zadna op data.
        else:
            pole_105_pa = 0
    # 310
    if "PZN310" in sklady:    
        sklad = obsolete_polozky_databaze.get(item).get("PZN310")
        # 310 IOH z I360
        # print(f'item {item} ma i360 qty na 310.')
        # projde linkty a nacete mnozstvi na skladu z nich jako ioh.
        for linka in sklad:
            ioh_310 = sklad.get(linka).get("Inventory on hand")
            ioh_310 = float(ioh_310.replace(",",""))
                # Planned available 310.
        op_310_itemu = order_plan_310.get(item)
        # Nektere itemy maji Inventory 360 data, ale nemaji order plan data (nefunkcni pegging) → tyto z kontroly vynechat.
        if op_310_itemu != None:
            # print(f'item {item} ma op data 310.')
            # Zkontrolovat STOCK ze neni divny odpovida ioh.
            for linka, data in op_310_itemu.items():
                order_type_310 = data.get("Order type txt")
                transaction_qty_310 = float(data.get("Quantity").replace(",",""))
                if order_type_310.upper() == "STOCK":
                    item_stock_op_310 = transaction_qty_310
                    pocet_stock_linek_itemu_310 += 1
            # Pokud stock linek je mene nez 2 je to OK → overit jeste, jestli si stock odpovida s IOH.
            # print(f'item {item} ma pocet STOC linek na op 310 {pocet_stock_linek_itemu_310}')
            if pocet_stock_linek_itemu_310 <2:
                # print(f'item {item} STOCK linky OK')
                # Kontrola zda si odpovida IOH z inventory dat a STOCK z op dat.
                # print(f'item {item} ma STOCK ioh z op 310 {item_stock_op_310} a ioh z I360 {ioh_310}')
                if ioh_310 == item_stock_op_310:
                    # print(f'item {item} IOH jsou stejne OK')
                    # Pokud ano, spocitat planned available na danem sklade.
                    pole_310_pa = planned_available_na_skladu(item, op_310_itemu)
                    # print(f'item {item} planned available op 310 {pole_310_pa}')
                else:
                    pole_310_pa = f'POZOR: item {item} ma jinou hodnotu STOCK transakce v pegging datech ({item_stock_op_310}) a v inventory datech ({ioh_310})!'                        
            else:                  
                pole_310_pa = f'POZOR! Pocet STOCK linek itemu {item} = {pocet_stock_linek_itemu_310}. Nebude kontrolovan.'           
        else:
            pole_310_pa = f'POZOR! {item} ma warehouse data 310, ale v order planu 310 nic nema (rozbity pegging u itemu). Nebude kontrolovan.'               
    # print(f'item {item} i360 qty na 310 = {ioh_310}.')                
    # Pokud item ma nejake QTY 0 ve whs 310 v i360, ale pri tom ma op stock vetsi nez jedna, je to diven.
    else:
        ioh_310 = 0
        op_310_itemu = order_plan_310.get(item)
        # Nektere itemy maji Inventory 360 data, ale nemaji order plan data (nefunkcni pegging) → tyto z kontroly vynechat.
        if op_310_itemu != None:
            # print(f'item {item} ma op data 310.')
            # Zkontrolovat STOCK ze neni divny odpovida ioh.
            pole_310_pa = planned_available_na_skladu(item, op_310_itemu)  
        # Pokud neni nic na skladu PZN 310 I360 LN ani nejsou zadna op data.
        else:
            pole_310_pa = 0
    
    # Prochazeni 100 na real demand + # Prochazeni 100 na sumu already na ceste Purcahse order.
    op_100_itemu = order_plan_100.get(item)
    if op_100_itemu != None:
        # Zkontrolovat STOCK ze neni divny.
        for linka, data in op_100_itemu.items():
            order_type_100 = data.get("Order type txt")
            transaction_qty_100 = float(data.get("Quantity").replace(",",""))
            if order_type_100.upper() == "STOCK":
                item_stock_op_100 = transaction_qty_100
                pocet_stock_linek_itemu_100 += 1
        # Pokud stock linek je mene nez 2 je to OK
        if pocet_stock_linek_itemu_100 <2:
            # Pokud ano, spocitat planned available na danem sklade.
            pole_100_real_demand = realny_demand_na_skladu(item, op_100_itemu)
            # Pokud ano, take spocitat sumu already na ceste Purcahse order.
            pole_100_pur_o_in_process = po_in_process_skladu(item, op_100_itemu)                     
        else:                  
            pole_100_real_demand = f'POZOR! Pocet STOCK linek itemu {item} = {pocet_stock_linek_itemu_100}. Real deman NEVEROHODNY!.'
            pole_100_pur_o_in_process = f'POZOR! Pocet STOCK linek itemu {item} = {pocet_stock_linek_itemu_100}. Real deman NEVEROHODNY!.'        
    else:
        pole_100_real_demand = f'Item {item} nema zadna data v order planu PZN100 → Real demand PZN100 = 0.'
        pole_100_pur_o_in_process = f'Item {item} nema zadna data v order planu PZN100 → Sum qty PZN100 = 0.'   
    
    # Prochazeni last transaction date 105 a 310.
    if item in obsolete_last_transactions_databaze:
        # 105
        if "PZN105" in obsolete_last_transactions_databaze.get(item):
            last_trans_date_105 = datetime.date(1,1,1)
            for linka in obsolete_last_transactions_databaze.get(item).get("PZN105"):
                transaction_date_105 = obsolete_last_transactions_databaze.get(item).get("PZN105").get(linka).get("Date")

                if transaction_date_105 > last_trans_date_105:
                    last_trans_date_105 = transaction_date_105
            pole_105_last_t = last_trans_date_105                
        else:
            pole_105_last_t = f'POZOR! item {item} nema zadne transtaction data.'
        # 310
        if "PZN310" in obsolete_last_transactions_databaze.get(item):
            last_trans_date_310 = datetime.date(1,1,1)
            for linka in obsolete_last_transactions_databaze.get(item).get("PZN310"):
                transaction_date_310 = obsolete_last_transactions_databaze.get(item).get("PZN310").get(linka).get("Date")

                if transaction_date_310 > last_trans_date_310:
                    last_trans_date_310 = transaction_date_310
            pole_310_last_t = last_trans_date_310                
        else:
            pole_310_last_t = f'POZOR! item {item} nema zadne transtaction data.'
    else:
        pole_105_last_t = f'POZOR! item {item} nema zadne transtaction data.'
        pole_310_last_t = f'POZOR! item {item} nema zadne transtaction data.'

    # Prochazeni kusovníků bud z SFExBFE databáze pro zjisteni, ve kterých všech boudách se item nachází.
    v_boudach = dotaz_pn_program(item, databaze_kusovniku_pro_dotaz, kvp_programy_pro_dotaz)
    v_boudach = v_boudach[1]
    v_boudach = [bouda_program.split("(")[0] for bouda_program in v_boudach]
    v_boudach = set(v_boudach)
    
    # Projiti databaze psr status databaze, jestli nektere z boud, ve kterych se nachazi jsou retired. 
    pocet_v_letajicich_boudach_psr = 0
    pocet_v_retired_boudach_psr = 0
    
    v_retired_boudach = list()
    
    # Projde postupne kazdou ukikatni boudu, ve ktere se PN nahchazi.
    for bouda in v_boudach:
        if bouda in psr_status_databaze:
            for linka in psr_status_databaze.get(bouda):
                status = psr_status_databaze.get(bouda).get(linka).get("Status")
                
                if "RETIRED" in status.upper():
                    v_retired_boudach.append(bouda)
                    pocet_v_retired_boudach_psr += 1
                else:
                    pocet_v_letajicich_boudach_psr +=1

    # Rozdeleni na zaklade podminek, co s tim dal.
    action_to_take_105 = "reset hodnota"
    action_to_take_310 = "reset hodnota"

    # "kose A,B,C"
    
    # A. Pokud je na sklade 105/310 IOH > 0: Skupinu A rovnou prevest na PZN
        # • Pokud je item bez alokace (nema v op zadne - (demandove) linky.)
            # • Pokud je na PZN 100 realny demand na item 
                # →→→ rovnou prevest vsechno mnozstvi na PZN100.
            # • Jinak:
                # Rozdelit na 3 skupiny podle stari last transakce na danem skladu. → dale projit.
                    # a) zadna transakce
                    # b) transakce v poslednich 3 letech
                    # c) transakce v poslednich 3 letech
    # B.  • Jinak nechat tam a zatim neresit. 

    # A. 105 
    # print(f'IOH 105 {ioh_105}, {type(ioh_105)}')
    # print(f'PA 105 {pole_105_pa}, {type(pole_105_pa)}')
    # print(f'real demand 100 {pole_100_real_demand}, {type(pole_100_real_demand)}')
    # print(f'last transaction 105 {pole_105_last_t}, {type(pole_105_last_t)}')

    # Pokud je na sklade 105/310 IOH > 0:
    if type(ioh_105) == float:
        if ioh_105 > 0:
            # Pokud je item bez alokace (nema v op zadne - (demandove) linky.)
            if type(pole_105_pa) == float:
                if ioh_105 == pole_105_pa:
                    # Pokud je na PZN 100 realny demand na item
                    if type(pole_100_real_demand) == float and pole_100_real_demand > 0:
                        action_to_take_105 = "prevest vse na PZN 100"
                            
                    # Jinak rozdelit na 3 skupiny podle stari last transakce na danem skladu. → dale projit.
                    else:
                        if type(pole_105_last_t) == datetime.date:
                            if pole_105_last_t > (datetime.date.today() - datetime.timedelta(365)):
                                action_to_take_105 = "dale proverit - last transakce v poslednim 1 roce."
                            elif pole_105_last_t > (datetime.date.today() - datetime.timedelta(1095)):
                                action_to_take_105 = "dale proverit - last transakce v poslednich 3 letech."
                            elif pole_105_last_t > (datetime.date.today() - datetime.timedelta(1825)):
                                action_to_take_105 = "dale proverit - last transakce v poslednich 5 letech."
                            else:
                                action_to_take_105 = "dale proverit - last transakce pred vice nez 5 lety."
                        else:
                            action_to_take_105 = "dale proverit - volne mnozstvi bez alokace a bez demandu na PZN100."        
                else:
                    action_to_take_105 = "Item ma nejake Demandove linky na skladu - nechat na skladu."                                
            else:
                action_to_take_105 = "Nereseno - neplatne mnozstvi Planned available 105."        
    else:
        action_to_take_105 = "Nereseno - neplatne mnozstvi IOH 105."

    # A. 310 
    # print(f'IOH 310 {ioh_310}, {type(ioh_310)}')
    # print(f'PA 310 {pole_310_pa}, {type(pole_310_pa)}')
    # print(f'real demand 100 {pole_100_real_demand}, {type(pole_100_real_demand)}')
    # print(f'last transaction 310 {pole_310_last_t}, {type(pole_310_last_t)}')

    # Pokud je na sklade 310/310 IOH > 0:
    if type(ioh_310) == float:
        if ioh_310 > 0:
            # Pokud je item bez alokace (nema v op zadne - (demandove) linky.)
            if type(pole_310_pa) == float:
                if ioh_310 == pole_310_pa:
                    # Pokud je na PZN 100 realny demand na item
                    if type(pole_100_real_demand) == float and pole_100_real_demand > 0:
                        action_to_take_310 = "prevest vse na PZN 100"
                            
                    # Jinak rozdelit na 3 skupiny podle stari last transakce na danem skladu. → dale projit.
                    else:
                        if type(pole_310_last_t) == datetime.date:
                            if pole_310_last_t > (datetime.date.today() - datetime.timedelta(365)):
                                action_to_take_310 = "dale proverit - last transakce v poslednim 1 roce."
                            elif pole_310_last_t > (datetime.date.today() - datetime.timedelta(1095)):
                                action_to_take_310 = "dale proverit - last transakce v poslednich 3 letech."
                            elif pole_310_last_t > (datetime.date.today() - datetime.timedelta(1825)):
                                action_to_take_310 = "dale proverit - last transakce v poslednich 5 letech."
                            else:
                                action_to_take_310 = "dale proverit - last transakce pred vice nez 5 lety."
                        else:
                            action_to_take_310 = "dale proverit - volne mnozstvi bez alokace a bez demandu na PZN100."        
                else:
                    action_to_take_310 = "Item ma nejake Demandove linky na skladu - nechat na skladu."
            else:
                action_to_take_310 = "Nereseno - neplatne mnozstvi Planned available 310."        
    else:        
        action_to_take_310 = "Nereseno - neplatne mnozstvi IOH 310."

    # tisk dat
    if type(ioh_105) == float:
        ioh_105 = str(ioh_105).replace(".",",")
    if type(ioh_310) == float:
        ioh_310 = str(ioh_310).replace(".",",")
    if type(pole_105_pa) == float:
        pole_105_pa = str(pole_105_pa).replace(".",",")
    if type(pole_310_pa) == float:
        pole_310_pa = str(pole_310_pa).replace(".",",")        
    if type(pole_100_real_demand) == float:
        pole_100_real_demand = str(pole_100_real_demand).replace(".",",") 
    if type(pole_100_pur_o_in_process) == float:
        pole_100_pur_o_in_process = str(pole_100_pur_o_in_process).replace(".",",")
    if len(v_boudach) == 0:
        v_boudach = "Neni v zadnem kusovniku v databazi"
    else:
        v_boudach = ",".join(v_boudach)

    if len(v_retired_boudach) == 0:
        v_retired_boudach = "PN Neni v zadne retired boude."        
    else:
        v_retired_boudach = ",".join(v_retired_boudach)
    if type(pocet_v_letajicich_boudach_psr) == float:
        pocet_v_letajicich_boudach_psr = str(pocet_v_letajicich_boudach_psr).replace(".",",")
    if type(pocet_v_retired_boudach_psr) == float:
        pocet_v_retired_boudach_psr = str(pocet_v_retired_boudach_psr).replace(".",",")    



    print(f'{item}|{ioh_105}|{pole_105_pa}|{pole_105_last_t}|{ioh_310}|{pole_310_pa}|{pole_310_last_t}|{pole_100_real_demand}|{pole_100_pur_o_in_process}|{v_boudach}|{pocet_v_letajicich_boudach_psr}|{v_retired_boudach}|{pocet_v_retired_boudach_psr}|{action_to_take_105}|{action_to_take_310}')

    with open("output.txt", "a", encoding="utf-8") as output:
        output.write(f'{item}|{ioh_105}|{pole_105_pa}|{pole_105_last_t}|{ioh_310}|{pole_310_pa}|{pole_310_last_t}|{pole_100_real_demand}|{pole_100_pur_o_in_process}|{v_boudach}|{pocet_v_letajicich_boudach_psr}|{v_retired_boudach}|{pocet_v_retired_boudach_psr}|{action_to_take_105}|{action_to_take_310}')
        output.write('\n')

print("\n\n--- %s seconds ---" % (time.time() - start_time))
input(f'\n\nKONEC PROGRAMU.\n\nVysledek je ulozen v souboru output.txt ve slozce Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\Obsolete analýza\n\nUkonci stiskem klavesy ENTER . . .')