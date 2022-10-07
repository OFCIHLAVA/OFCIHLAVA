from funkce.cq_data import data_date_formating, data_import, data_headings,import_data_cleaning, order_plan_database_pzn100, order_plan_database_pzn105, order_plan_database_pzn310, obsolete_polozky_database
from funkce.funkce_prace import dotaz_pn_program, nacteni_databaze_boud_pro_dotaz, programy_boud
from funkce.prevody_dotazy import planned_available_na_skladu, po_in_process_skladu, realny_demand_na_skladu

import datetime
import sys
import time


# Priprava databazi pro dotazovani v jakych boudach je item obsazen.
path_kusovniky_databaze = 'Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\SFExBFExMIX\\databaze\\databaze boud s kusovniky.txt'
path_program_databaze = 'Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\SFExBFExMIX\\databaze\\seznam programu.txt'

databaze_kusovniku_pro_dotaz = nacteni_databaze_boud_pro_dotaz(path_kusovniky_databaze)
seznam_boud_z_databaze_kusovniku = [key for key in databaze_kusovniku_pro_dotaz]
print("Databaze bud s kusovniky nactena a pripravena pro dotazovani . . .")

kvp_programy_pro_dotaz = programy_boud(path_program_databaze)
seznam_boud_z_databaze_programu = [key for key in kvp_programy_pro_dotaz]
print("Databaze programu vsech boud nacetenaa pripravena pro dotazovani . . .\n")

obsolete_file_import = "Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\Obsolete analýza\\polozky_proverit\\obsolete polozky.txt"
obsolete_order_plan_import = "Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\Obsolete analýza\\order_plany\\order_plan_100_105_310.txt"
obsolete_last_transactions_import = "Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\Obsolete analýza\\posledni_transakce\\obselete_last_transaction.txt"

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
print(f'Seznam OBSOLETE polozek k provereni:\n')
for item in obs_polozky_seznam:
    print(item)

print(f'\nCelkem polozek k provereni: {len(obs_polozky_seznam)}.')    

### LAST TRANSACTIONS
input(f'\nKROK 2:\n\nVloz seznam polozek vyse do reportu:\nusers/Ondrej.Rott/Obsolete analyza/Issued for spares.eq\na vysledek reportu uloz d slozky:/nY:/Departments/Sales and Marketing/Aftersales/11_PLANNING/23_Python_utilities/Obsolete analýza/posledni_transakce pod nazvem:\nobselete_last_transaction.txt\nPote pokracuj stiskem ENTER . . . ')

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


# for item, sklady in obsolete_last_transactions_databaze.items():
#     print(f'{item}')
#     for sklad, linky in sklady.items():
#         print(sklad)
#         print(len(linky))
#         # for linka, data in linky.items():
#         #     print(f'{linka} :  {data}')
# 

### ORDER PLANY

input(f'\nKROK 3:\n\nVloz seznam polozek vyse do reportu:\nusers/Ondrej.Rott/Obsolete analyza/PZN100+105+310_order_plan.eq\na vysledek reportu uloz d slozky:/nY:/Departments/Sales and Marketing/Aftersales/11_PLANNING/23_Python_utilities/Obsolete analýza/order_plany pod nazvem:\norder_plan_100_105_310.txt\nPote pokracuj stiskem ENTER . . . ')

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

input(f'\n\nProgram ready.\nZabere to cca {round(len(obs_polozky_seznam)*0.141)} sekund / {round(len(obs_polozky_seznam)*0.141/60,1)} minut\n\nSpust stiskem ENTER . . .')
start_time = time.time()

print(f'Item|planned available 105 (pegging)|Last transaction 105|planned available 310 (pegging)|Last transaction 310|Real demand 100 (pegging)|Sum qty already incoming Purchase orders na 100')
with open("output.txt", "w", encoding="utf-8") as output:
    output.write(f'Item|planned available 105 (pegging)|Last transaction 105|planned available 310 (pegging)|Last transaction 310|Real demand 100 (pegging)|Sum qty already incoming Purchase orders na 100|Obsazeno v boudach\n')
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
            for linka in obsolete_last_transactions_databaze.get(item).get("PZN105"):
                last_trans_date_105 = datetime.date(1,1,1)
                transaction_date_105 = obsolete_last_transactions_databaze.get(item).get("PZN105").get(linka).get("Date")

                if transaction_date_105 > last_trans_date_105:
                    last_trans_date_105 = transaction_date_105
            pole_105_last_t = last_trans_date_105                
        else:
            pole_105_last_t = f'POZOR! item {item} nema zadne transtaction data.'
        # 310
        if "PZN310" in obsolete_last_transactions_databaze.get(item):
            for linka in obsolete_last_transactions_databaze.get(item).get("PZN310"):
                last_trans_date_310 = datetime.date(1,1,1)
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
    # tisk dat

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




    print(f'{item}|{pole_105_pa}|{pole_105_last_t}|{pole_310_pa}|{pole_310_last_t}|{pole_100_real_demand}|{pole_100_pur_o_in_process}|{v_boudach}')

    with open("output.txt", "a", encoding="utf-8") as output:
        output.write(f'{item}|{pole_105_pa}|{pole_105_last_t}|{pole_310_pa}|{pole_310_last_t}|{pole_100_real_demand}|{pole_100_pur_o_in_process}|{v_boudach}')
        output.write('\n')

print("\n\n--- %s seconds ---" % (time.time() - start_time))
input(f'\n\nKONEC PROGRAMU.\n\nVysledek je ulozen v souboru output.txt ve slozce Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\Obsolete analýza\n\nUkonci stiskem klavesy ENTER . . .')