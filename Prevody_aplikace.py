import warnings
import openpyxl as excel
import datetime
from turtle import heading
from urllib.request import DataHandler
from prevody_funkce import prevody_dotazy
from prevody_data import cq_data, excel_data
import _warnings

warnings.simplefilter(action='ignore', category=UserWarning)

#########
# INPUT:
#########

print(f'\nNejprve je potreba ziskat data prodeju z Master planu.')
print(f'1.KROK: Spust CQ report "TEST_Sales_order_lines_master_plan.eq" ve slozce "webreports\\After Sales\\Ondra test" a vysledek reportu exportuj / uloz jako "master plan.txt" soubor do slozky Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\Převody\\Master plan txt\\...\n')
input(f'Az bude ulozeno ve slozce, pokracuj stiskem klavesy ENTER...')

# Zadani poctu dni jak dlouho dopredu proverovat prodeje.
kal_dnu_k_provereni_ode_dneska = ""
while type(kal_dnu_k_provereni_ode_dneska) != int or kal_dnu_k_provereni_ode_dneska < 0:
    kal_dnu_k_provereni_ode_dneska = input(f'\n2.KROK: Zadej kolik kalendarnich dni ode dneska chces proverit shortage linky (0 proveri pouze dnesek): ')    
    try:
        kal_dnu_k_provereni_ode_dneska = int(kal_dnu_k_provereni_ode_dneska)
        if kal_dnu_k_provereni_ode_dneska < 0:
            print(f'Je potreba zadat cele nezaporne cislo. Zkus znovu...')
    except ValueError:
        print(f'Je potreba zadat cele nezaporne cislo. Zkus znovu...')

# Priprava datumu, jak dlouho dopredu se maji proverovat linky Master planu.
proverit_do_datumu = excel_data.do_datumu_proverit_master_plan(kal_dnu_k_provereni_ode_dneska)
print(f'\nBudou se proverovat linky z Master Planu s datumy od dneska do {proverit_do_datumu.strftime("%d/%m/%Y").replace("/", ".")}.\n')

########
# Master plan data priprava.
########

# Nacteni Master planu CQ.
master_plan_data = cq_data.data_import("Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\Převody\\Master plan txt\\master plan.txt")
print(f'Master plan data nactena...\n')

# Vytvoreni zahlavi master planu CQ.
master_plan_zahlavi = cq_data.data_headings_master_plan(master_plan_data)
print(f'Zahlavi Master planu vytvoreno...\n')
# print(master_plan_zahlavi)

# Ocisteni CQ dat pro dalsi zpracovani.
master_plan_data = cq_data.import_data_cleaning(master_plan_data)
print(f'Master plan data ocistena...\n')

# Pridani sloupce dat s row linky Master planu do Master plannu dat.
master_plan_data = cq_data.add_row_number_to_master_plan_data(master_plan_data)
print(f'Pridan udaj o row linky z Master planu do dat Master planu na prvni pozici...\n')

# Pridani Availability sloupce na konec dat.
master_plan_data = cq_data.add_availability_master_plan_data(master_plan_data, master_plan_zahlavi)
print(f'Pridan udaj o Availibility do dat Master planu na posledni pozici...\n')

# Opraveni formatu datumu v Master plan datech.
master_plan_data = cq_data.data_date_formating(master_plan_data, master_plan_zahlavi)
print(f'Opraven format datumu v Master plan datech...\n')

##############################
# Vytvoreni seznamu itemu z Master planu ke vlozeni do CQ reportu order planu. (Pro jake itemy budeme proverovat order plany.)
##############################

# Vytvoreni seznamu itemu z Master planu k provereni podle zadaneho datumu od uzivatele.
seznam_itemu_pro_order_plany = cq_data.seznam_itemu_pro_order_plany_cq(master_plan_data, master_plan_zahlavi, proverit_do_datumu)
print(f'Vytvoren seznam itemu ke vlozeni do CQ reportu order planu 100+105...\n')

# Vytisteni seznamu itemu z Master planu k provereni pro zadani do CQ reportu order planu.
for item in seznam_itemu_pro_order_plany:
    print(item)
print(f'\n3.KROK: Seznam polozek vyse je potreba vlozit do CQ reportu "PZN100+105_order_plan.eq" ve slozce "webreports\\After Sales\\Ondra test" a vysledek reportu exportovat / ulozit jako "order plan 100+105.txt" soubor do slozky Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\Převody\\Order plan...\n')
input(f'Az bude ulozeno ve slozce, pokracuj stiskem klavesy ENTER...')  


# Puvodni verze nacitajici data z Master planu excelu.
#↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓

## # Nacteni Master planu.
## wb_master_plan = excel.load_workbook("Y:\\Departments\\Sales and Marketing\\Shared\\Aftersales\\SCCZ-AS-F003 Aftersales Master Plan.xlsm")
## sheet1_master_plan = wb_master_plan.worksheets[0]
## print(f'Master Plan nacten...\n')
## 
## # Vytvoreni zahlavi master planu.
## master_plan_zahlavi = excel_data.zahlavi_master_planu(sheet1_master_plan)
## print(f'Zahlavi Master planu vytvoreno...\n')
## # print(master_plan_zahlavi)
## 
## ##############################
## # Vytvoreni seznamu itemu z Master planu ke vlozeni do CQ reportu order planu. (Pro jake itemy budeme proverovat order plany.)
## ##############################
## 
## # Vytvoreni seznamu itemu z Master planu k provereni podle zadaneho datumu od uzivatele.
## seznam_itemu_pro_order_plany = excel_data.seznam_itemu_pro_order_plany_excel(sheet1_master_plan, master_plan_zahlavi, proverit_do_datumu)
## print(f'Vytvoren seznam itemu ke vlozeni do CQ reportu order planu 100+105...\n')

##############################
# Order plany CQ data priprava.
##############################

order_plan_data = cq_data.data_import("Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\Převody\\Order plan\\order plan 100+105.txt")
print(f'CQ data nactena...\n')

# Vutvoreni zahlavi z dat z CQ exportu.
order_plan_headings = cq_data.data_headings_order_plan(order_plan_data)
print(f'CQ zahlavi pripraveno...\n')
# print(order_plan_headings)

# Ocisteni CQ dat pro dalsi zpracovani.
order_plan_data = cq_data.import_data_cleaning(order_plan_data)
print(f'CQ data ocistena...\n')

# Pridani sloupce dat s ID linky do CQ dat.
order_plan_data = cq_data.add_line_id_to_order_plan_data(order_plan_data, order_plan_headings)
print(f'Pridan udaj o ID linky do dat z CQ na prvni pozici...\n')

# Opraveni formatu datumu v CQ datech.
order_plan_data = cq_data.data_date_formating(order_plan_data, order_plan_headings)
print(f'Opraven format datumu v CQ datech...\n')

# Vytvoreni Item order planu PZN100.
order_plan_pzn_100 = cq_data.order_plan_database_pzn100(order_plan_data, order_plan_headings)
print(f'Order plan PZN100 vytvoren...\n')
# print(order_plan_pzn_100)

# Vytvoreni Item order planu PZN105.
order_plan_pzn_105 = cq_data.order_plan_database_pzn105(order_plan_data, order_plan_headings)
print(f'Order plan PZN105 vytvoren...\n')
# print(order_plan_pzn_105)

##############################
# Vystup priprava.
##############################

# Priprava Zahlavi pro shortage linky.
vystup_zahlavi = cq_data.zahlavi_vystupu_cq(master_plan_zahlavi)
print(f'Zahlavi vystupu vytvoreno...\n')
# print(vystup_zahlavi)

# Puvodni verze z excel Master planu.
## # Priprava Zahlavi pro shortage linky.
## vystup_zahlavi = excel_data.zahlavi_vystupu_excel(sheet1_master_plan, master_plan_zahlavi)
## print(f'Zahlavi vystupu vytvoreno...\n')
## # print(vystup_zahlavi)

# Poskladani shortage linek z Master planu.
shortage_linky_proverit = cq_data.shortage_linky_master_planu_cq(master_plan_data, master_plan_zahlavi, proverit_do_datumu, order_plan_pzn_100, order_plan_pzn_105)
print(f'Shortage linky vytvoreny...\n')

# Puvodni verze z excel Master planu.
## # Poskladani shortage linek z Master planu.
## shortage_linky_proverit = excel_data.shortage_linky_master_planu_excel(sheet1_master_plan, master_plan_zahlavi, proverit_do_datumu, order_plan_pzn_100, order_plan_pzn_105)
## print(f'Shortage linky vytvoreny...\n')

# Doplneni SUM req Qty v dany den.
excel_data.doplneni_sum_ordered_qty_do_vystupu(shortage_linky_proverit, vystup_zahlavi)
print(f'\nSum of required Qty pridano do vystupu...\n')

# Doplneni Planned available na PZN105 na dane PDD linky.
prevody_dotazy.planned_available_na_skladu(shortage_linky_proverit, order_plan_pzn_105, vystup_zahlavi)
print(f'\nPlanned available PZN105 pridano do vystupu...\n')

# Doplneni Planned available na PZN100 na dane PDD linky.
prevody_dotazy.planned_available_na_skladu(shortage_linky_proverit, order_plan_pzn_100, vystup_zahlavi)
print(f'\nPlanned available PZN100 pridano do vystupu...\n')

# Nacteni Tabulky prevodu na PZN105.
wb_tabulka_prevodu = excel.load_workbook("Y:\\Departments\\Logistics\\Shared\WAREHOUSE\\After sales\\Prevody na PZN105_2022.xlsx")
sheet1_tabulka_prevodu = wb_tabulka_prevodu.worksheets[0]
print(f'Tabulka prevodu nactena...\n')

# Doplneni Already requested in tabulka prevodu PZN105.
excel_data.doplneni_already_zadano_do_vystupu(sheet1_tabulka_prevodu, shortage_linky_proverit, vystup_zahlavi, 5)
print(f'\nAlready requested in tabulka prevodu PZN105 pridano do vystupu...\n')

# Doplneni nejblizsiho datumu + Planned available qty, kdy na PZN105 bude Planned available alespon 0.
prevody_dotazy.next_planned_available_date_not_shortage_sklad(shortage_linky_proverit, order_plan_pzn_105, vystup_zahlavi)
print(f'\nNejblizsi Planned available PZN105 alespon 0 doplneno do vystupu...\n')

# Doplneni nejblizsiho datumu + Planned available qty, kdy na PZN100 bude Planned available alespon 0.
prevody_dotazy.next_planned_available_date_not_shortage_sklad(shortage_linky_proverit, order_plan_pzn_100, vystup_zahlavi)
print(f'\nNejblizsi Planned available PZN100 alespon 0 doplneno do vystupu...\n')

# Doplneni udaje, zda mozno prevest z PZN100 na PZN105 aniz by vznikl shortage na ostatnich linkach order planu PZN100.
prevody_dotazy.next_planned_available_date_simulate_prevody(shortage_linky_proverit, order_plan_pzn_105, order_plan_pzn_100, vystup_zahlavi)
print(f'\nInfo, zda mozno prevest doplneno do vystupu...\n')

# Doplneni zahlavi do seznamu shortage linek k provereni.
excel_data.doplneni_zahlavi_do_vystupu(shortage_linky_proverit, vystup_zahlavi)
print(f'\nZahlavi doplneno do vystupu...\n')

# Vytisteni vystupu po jednotlivych linkach.
for line in shortage_linky_proverit: # kontrolni TISK
    row_to_print = []
    for pole in line:     
        row_to_print.append(str(pole))
    i = 0
    for cislo in row_to_print[8:14]:
        row_to_print[8+i] = str(cislo).replace(".",",")
        i+=1
    print("|".join(row_to_print))

print(f'\nHOTOVO')
print(f'\n4.KROK: Blok dat vyse je potreba zkopirovat do souboru "kontrolovadlo prevodu.xlsx" ve slozce "Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\Převody\\kontrolovadlo prevodu.xlsx" a rozdelit text to columns podle znaku "|". Akce k provedeni je pak ve sloupci "Mozno prevest z PZN100 aniz by se ohrozily budouci linky na PZN100?".')

input(f'\nStiskni ENTER pro ukonceni programu...')