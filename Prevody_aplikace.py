import warnings
import openpyxl as excel
import datetime
from turtle import heading
from urllib.request import DataHandler
from prevody_funkce import prevody_dotazy
from prevody_data import cq_data, excel_data
import _warnings

warnings.simplefilter(action='ignore', category=UserWarning)
# INPUT: 
kal_dnu_k_provereni_ode_dneska = 14



### CQ data priprava.
order_plan_data = cq_data.data_import("C:\\Users\\Ondrej.rott\\Documents\\Python\\Prevody\\CQ exporty\\report.txt")
print(f'CQ data nactena...\n')

# Vutvoreni zahlavi z dat z CQ exportu.
order_plan_headings = cq_data.data_headings(order_plan_data)
print(f'CQ zahlavi pripraveno...\n')
print(order_plan_headings)

# Ocisteni CQ dat pro dalsi zpracovani.
order_plan_data = cq_data.import_data_cleaning(order_plan_data)
print(f'CQ data ocistena...\n')

# Pridani sloupce dat s ID linky do CQ dat.
order_plan_data = cq_data.add_line_id_to_data(order_plan_data, order_plan_headings)
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

########
# Master plan data priprava.
########

# Nacteni Master planu.
wb_master_plan = excel.load_workbook("Y:\\Departments\\Sales and Marketing\\Shared\\Aftersales\\SCCZ-AS-F003 Aftersales Master Plan.xlsm")
sheet1_master_plan = wb_master_plan.worksheets[0]
print(f'Master Plan nacten...\n')

# Nacteni Tabulky prevodu na PZN105.
wb_tabulka_prevodu = excel.load_workbook("Y:\\Departments\\Logistics\\Shared\WAREHOUSE\\After sales\\Prevody na PZN105_2022.xlsx")
sheet1_tabulka_prevodu = wb_tabulka_prevodu.worksheets[0]
print(f'Tabulka prevodu nactena...\n')

# Vytvoreni zahlavi master planu.
master_plan_zahlavi = excel_data.zahlavi_master_planu(sheet1_master_plan)
print(f'Zahlavi Master planu vytvoreno...\n')
print(master_plan_zahlavi)

# Priprava Zahlavi pro shortage linky.
vystup_zahlavi = excel_data.zahlavi_vystupu(sheet1_master_plan, master_plan_zahlavi)
print(f'Zahlavi vystupu vytvoreno...\n')
print(vystup_zahlavi)
print()

# Priprava datumu, jak dlouho dopredu se maji proverovat linky Master planu.
proverit_do_datumu = excel_data.do_datumu_proverit_master_plan(kal_dnu_k_provereni_ode_dneska)
print(f'Budou se proverovat linky z Master Planu s datumy na pristich {kal_dnu_k_provereni_ode_dneska} dnu.\n')

# Poskladani shortage linek z Master planu.
shortage_linky_proverit = excel_data.shortage_linky_master_planu(sheet1_master_plan, master_plan_zahlavi, proverit_do_datumu, order_plan_pzn_100, order_plan_pzn_105)
print(f'Shortage linky vytvoreny...\n')

# Doplneni SUM req Qty v dany den.
excel_data.doplneni_sum_ordered_qty_do_vystupu(shortage_linky_proverit, vystup_zahlavi)
print(f'\nSum of required Qty pridano do vystupu...\n')

# Doplneni Planned available na PZN105 na dane PDD linky.
prevody_dotazy.planned_available_na_skladu(shortage_linky_proverit, order_plan_pzn_105, vystup_zahlavi)
print(f'\nPlanned available PZN105 pridano do vystupu...\n')

# Doplneni Planned available na PZN100 na dane PDD linky.
prevody_dotazy.planned_available_na_skladu(shortage_linky_proverit, order_plan_pzn_100, vystup_zahlavi)
print(f'\nPlanned available PZN100 pridano do vystupu...\n')

# Doplneni Already requested in tabulka prevodu PZN105.
excel_data.doplneni_already_zadano_do_vystupu(sheet1_tabulka_prevodu, shortage_linky_proverit, vystup_zahlavi, 5)
print(f'\nAlready requested in tabulka prevodu PZN105 pridano do vystupu...\n')

# Doplneni nejblizsiho datumu + Planned available qty, kdy na PZN105 bude Planned available alespon 0.
prevody_dotazy.next_planned_available_date_not_shortage_sklad(shortage_linky_proverit, order_plan_pzn_105, vystup_zahlavi)
print(f'\nNejblizsi Planned available PZN105 alespon 0 doplneno do vystupu...\n')

# Doplneni nejblizsiho datumu + Planned available qty, kdy na PZN100 bude Planned available alespon 0.
prevody_dotazy.next_planned_available_date_not_shortage_sklad(shortage_linky_proverit, order_plan_pzn_100, vystup_zahlavi)
print(f'\nNejblizsi Planned available PZN100 alespon 0 doplneno do vystupu...\n')

for line in shortage_linky_proverit: # kontrolni TISK
    row_to_print = []
    for pole in line:     
        row_to_print.append(str(pole))
    i = 0
    for cislo in row_to_print[8:14]:
        row_to_print[8+i] = str(cislo).replace(".",",")
        i+=1
    print("|".join(row_to_print))

# Doplneni udaje, zda mozno prevest z PZN100 na PZN105 aniz by vznikl shortage na ostatnich linkach order planu PZN100.
prevody_dotazy.next_planned_available_date_simulate_prevody(shortage_linky_proverit, order_plan_pzn_105, order_plan_pzn_100, vystup_zahlavi)
print(f'\nInfo, zda mozno prevest doplneno do vystupu...\n')

# Doplneni zahlavi do seznamu shortage linek k provereni.
excel_data.doplneni_zahlavi_do_vystupu(shortage_linky_proverit, vystup_zahlavi)

for line in shortage_linky_proverit: # kontrolni TISK
    row_to_print = []
    for pole in line:     
        row_to_print.append(str(pole))
    i = 0
    for cislo in row_to_print[8:14]:
        row_to_print[8+i] = str(cislo).replace(".",",")
        i+=1
    print("|".join(row_to_print))

print(f'Hotovo')
