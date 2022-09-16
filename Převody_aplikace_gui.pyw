from tkinter import *
import datetime
import openpyxl as excel

from prevody_funkce import prevody_dotazy
from prevody_data import cq_data, excel_data

def run_program():

    if 1 in stage:    
        log_window['state'] = 'normal'
        
        run_btn_text.set("CONTINUE")
        log_window.delete(1.0, END)
        log_text_0 = """1.KROK:\nZadej počet dní k prověření a spusť CQ report "TEST_Sales_order_lines_master_plan.eq" ve složce: "webreports\After Sales\Ondra test".\nVýsledek reportu exportuj/ulož jako "master plan.txt" soubor do složky\nY:\Departments\Sales and Marketing\Aftersales\11_PLANNING\23_Python_utilities\Převody\Master plan txt\n\naž bude připraveno, pokračuj tlačítkem CONTINUE...\n"""
        log_window.insert(1.0, log_text_0)
        stage.pop()
        stage.append(2)

        log_window['state'] = 'disabled'
    elif 2 in stage:       
        
        ##############################
        # Master plan data CQ priprava.
        ##############################
        
        log_window['state'] = 'normal'
        items_to_order_plans['state'] = 'normal'

        run_btn_text.set("VÝSLEDEK")
        
        kal_dnu_k_provereni_ode_dneska = int(next_cal_days_to_check.get())
        
        global proverit_do_datumu                
        proverit_do_datumu = excel_data.do_datumu_proverit_master_plan(kal_dnu_k_provereni_ode_dneska)
        do_datumu.set(proverit_do_datumu.strftime("%d/%m/%Y").replace("/", "."))        
        log_text_1 = f'2.KROK:\n• Budou se prověřovat linky z Master Plánů s datumy od dneška do {proverit_do_datumu.strftime("%d/%m/%Y").replace("/", ".")} . . .\n'
        log_window.delete(1.0, END)
        log_window.insert(END, log_text_1)
        
        # Nacteni Master planu CQ.
        global master_plan_data        
        master_plan_data = cq_data.data_import("Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\Prevody_PZN100_PZN105\\Master_plan_txt\\Prevody_Sales_order_lines_master_plan.cq")
        log_text_2 = f'• Master plan data načtena . . .\n' 
        log_window.insert(END, log_text_2)

        # Vytvoreni zahlavi master planu CQ.
        global master_plan_zahlavi
        master_plan_zahlavi = cq_data.data_headings_master_plan(master_plan_data)
        log_text_3 = f'• Záhlavi Master planu vytvořeno . . .\n' 
        log_window.insert(END, log_text_3)

        # Ocisteni CQ dat pro dalsi zpracovani.
        master_plan_data = cq_data.import_data_cleaning(master_plan_data)
        log_text_4 = f'• Master plan data očištěna . . .\n'
        log_window.insert(END, log_text_4)

        # Pridani sloupce dat s row linky Master planu do Master plannu dat.
        master_plan_data = cq_data.add_row_number_to_master_plan_data(master_plan_data)
        log_text_5 = f'• Přidán údaj o row linky z Master planu do CQ dat Master planu na první pozici . . .\n'
        log_window.insert(END, log_text_5)

        # Pridani Availability sloupce na konec dat.
        master_plan_data = cq_data.add_availability_master_plan_data(master_plan_data, master_plan_zahlavi)
        log_text_6 = f'• Přidán údaj o Availability do CQ dat Master planu na poslední pozici . . .\n'
        log_window.insert(END, log_text_6)

        # Opraveni formatu datumu v Master plan datech.
        master_plan_data = cq_data.data_date_formating(master_plan_data, master_plan_zahlavi)
        log_text_7 = f'• Opraven formát datumu v CQ Master plan datech . . .\n'
        log_window.insert(END, log_text_7)

        ##############################
        # Vytvoreni seznamu itemu z Master planu ke vlozeni do CQ reportu order planu. (Pro jake itemy budeme proverovat order plany.)
        ##############################
        # Vytvoreni seznamu itemu z Master planu k provereni podle zadaneho datumu od uzivatele.
        seznam_itemu_pro_order_plany = cq_data.seznam_itemu_pro_order_plany_cq(master_plan_data, master_plan_zahlavi, proverit_do_datumu)
        log_text_8 = f'• Vytvořen seznam itemů ke vložení do CQ reportu order plánů 100+105 . . .\n'
        log_window.insert(END, log_text_8)        

        # Vytisteni seznamu itemu z Master planu k provereni pro zadani do CQ reportu order planu.
        for item in seznam_itemu_pro_order_plany:
            items_to_order_plans.insert(END,f'{item}\n')

        log_text_9 = """\n3. KROK:\n• Seznam itemů vlevo je třeba zadat do CQ reportu "PZN100+105_order_plan.eq"\nve složce "webreports\After Sales\Ondra test" a výsledek exportuj/ulož jako\n"order plan 100+105.txt" soubor do složky\nY:\Departments\Sales and Marketing\Aftersales\11_PLANNING\\n23_Python_utilities\Převody\Order plan...\n\naž bude připraveno, pokračuj tlačítkem CONTINUE...\n"""
        log_window.insert(END, log_text_9)  

        stage.pop()
        stage.append(3)

        log_window['state'] = 'disabled'
        items_to_order_plans['state'] = 'disabled'

    elif 3 in stage:
        ##############################
        # Order plany CQ data priprava.
        ##############################
        
        log_window['state'] = 'normal'
        output['state'] = 'normal'

        # Smazani dat jestli nejaka jasou v outputu.
        output.delete(1.0, END)
        
        global order_plan_data
        order_plan_data = cq_data.data_import("Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\Prevody_PZN100_PZN105\\Order plan\\order plan 100+105.txt")
        log_text_10 = f'\n• CQ data načtena . . .\n' 
        log_window.insert(END, log_text_10)               

        # Vytvoreni zahlavi z dat z CQ exportu.
        global order_plan_headings
        order_plan_headings = cq_data.data_headings_order_plan(order_plan_data)
        log_text_11 = f'• CQ záhlaví připraveno . . .\n' 
        log_window.insert(END, log_text_11)

        # Ocisteni CQ dat pro dalsi zpracovani.
        order_plan_data = cq_data.import_data_cleaning(order_plan_data)
        log_text_12 = f'• CQ data očištěna . . .\n' 
        log_window.insert(END, log_text_12)        

        # Pridani sloupce dat s ID linky do CQ dat.
        order_plan_data = cq_data.add_line_id_to_order_plan_data(order_plan_data, order_plan_headings)
        log_text_13 = f'• Přidán údaj o ID linky do dat z CQ na první pozici . . .\n' 
        log_window.insert(END, log_text_13)            

        # Opraveni formatu datumu v CQ datech.
        order_plan_data = cq_data.data_date_formating(order_plan_data, order_plan_headings)
        log_text_14 = f'• Opraven formát datumu v CQ datech . . .\n' 
        log_window.insert(END, log_text_14)          

        # Vytvoreni Item order planu PZN100.
        global order_plan_pzn_100
        order_plan_pzn_100 = cq_data.order_plan_database_pzn100(order_plan_data, order_plan_headings)
        log_text_15 = f'• Order pan PZN100 vytvořen . . .\n' 
        log_window.insert(END, log_text_15)

        # Vytvoreni Item order planu PZN105.
        global order_plan_pzn_105
        order_plan_pzn_105 = cq_data.order_plan_database_pzn105(order_plan_data, order_plan_headings)
        log_text_15 = f'• Order pan PZN100 vytvořen . . .\n\n\n'        
        log_window.insert(END, log_text_15)

        ##############################
        # Vystup priprava.
        ##############################
        # Priprava Zahlavi pro shortage linky.
        global vystup_zahlavi
        vystup_zahlavi = cq_data.zahlavi_vystupu_cq(master_plan_zahlavi)
        log_text_16 = f'• Záhlaví výstupu vytvořeno . . .\n'         
        log_window.insert(END, log_text_16)        

        # Poskladani shortage linek z Master planu.
        global shortage_linky_proverit
        shortage_linky_proverit = cq_data.shortage_linky_master_planu_cq(master_plan_data, master_plan_zahlavi, proverit_do_datumu, order_plan_pzn_100, order_plan_pzn_105)
        log_text_17 = f'• Shortage linky vytvořeny . . .\n'         
        log_window.insert(END, log_text_17)         

        # Doplneni SUM req Qty v dany den.
        excel_data.doplneni_sum_ordered_qty_do_vystupu(shortage_linky_proverit, vystup_zahlavi)
        log_text_18 = f'• Sum of required Qty přidáno do výstupu . . .\n'         
        log_window.insert(END, log_text_18)  

        # Doplneni Planned available na PZN105 na dane PDD linky + vytvoreni seznamu purchase orders , ktere uz se pocitaji do Planned available, ale jeste musi dnes prijit.
        prevody_dotazy.planned_available_na_skladu(shortage_linky_proverit, order_plan_pzn_105, vystup_zahlavi)
        log_text_19 = f'• Planned available na PZN105 přidáno do výstupu . . .\n'         
        log_window.insert(END, log_text_19)

        # Doplneni Planned available na PZN100 na dane PDD linky.
        prevody_dotazy.planned_available_na_skladu(shortage_linky_proverit, order_plan_pzn_100, vystup_zahlavi)
        log_text_20 = f'• Planned available na PZN100 přidáno do výstupu . . .\n'         
        log_window.insert(END, log_text_20) 

        # Doplneni Already requested in tabulka prevodu PZN105. (Pro verzi bez Excel modulu neni zatim k dispozici)

        # Nacteni Tabulky prevodu na PZN105.
        wb_tabulka_prevodu = excel.load_workbook("Y:\\Departments\\Logistics\\Shared\WAREHOUSE\\After sales\\Prevody na PZN105_2022.xlsx")
        sheet1_tabulka_prevodu = wb_tabulka_prevodu.worksheets[0]
        log_text_21 = f'• Tabulka prevodu nactena . . .\n'         
        log_window.insert(END, log_text_21)

        excel_data.doplneni_already_zadano_do_vystupu(sheet1_tabulka_prevodu, shortage_linky_proverit, vystup_zahlavi, 5)
        log_text_22 = f'• lready requested in tabulka prevodu PZN105 pridano do vystupu . . .\n'         
        log_window.insert(END, log_text_22)

        ## Pro verzi bez excel modulu nize:
        #for line in shortage_linky_proverit:
        #    line.append("Nelze zjistit bez Excel modulu")
        # log_text_21 = f'• Already requsted v tabulce převodů NEBYLO doplněno . . .\n' 

        # Doplneni nejblizsiho datumu + Planned available qty, kdy na PZN105 bude Planned available alespon 0.
        prevody_dotazy.next_planned_available_date_not_shortage_sklad(shortage_linky_proverit, order_plan_pzn_105, vystup_zahlavi)
        log_text_23 = f'• Nejbližší Planned available PZN105 alespoň 0 přidáno do výstupu . . .\n'         
        log_window.insert(END, log_text_23)         

        # Doplneni nejblizsiho datumu + Planned available qty, kdy na PZN100 bude Planned available alespon 0.
        prevody_dotazy.next_planned_available_date_not_shortage_sklad(shortage_linky_proverit, order_plan_pzn_100, vystup_zahlavi)
        log_text_24 = f'• Nejbližší Planned available PZN100 alespoň 0 přidáno do výstupu . . .\n'         
        log_window.insert(END, log_text_24)  

        # Doplneni udaje, zda mozno prevest z PZN100 na PZN105 aniz by vznikl shortage na ostatnich linkach order planu PZN100.
        prevody_dotazy.next_planned_available_date_simulate_prevody(shortage_linky_proverit, order_plan_pzn_105, order_plan_pzn_100, vystup_zahlavi)
        log_text_25 = f'• Info zda možno převést přidáno do výstupu . . .\n'         
        log_window.insert(END, log_text_25)

        # Doplneni zahlavi do seznamu shortage linek k provereni.
        excel_data.doplneni_zahlavi_do_vystupu(shortage_linky_proverit, vystup_zahlavi)
        log_text_26 = f'• Záhlaví přidáno do výstupu . . .\n'         
        log_window.insert(END, log_text_26)

        log_text_27 = f'\n• HOTOVO . . .\n\nVýsledek je třeba zkopírovat do excelu a rozdělit TEXT CO COLUMNS podle zanku "|".\n'         
        log_window.insert(END, log_text_27)        


        # Vytisteni vystupu po jednotlivych linkach.
        for line in shortage_linky_proverit: # kontrolni TISK
            row_to_print = []
            for pole in line:     
                row_to_print.append(str(pole))
            i = 0
            for cislo in row_to_print[8:14]:
                row_to_print[8+i] = str(cislo).replace(".",",")
                i+=1
            o="|".join(row_to_print)
            output.insert(END, o)
            output.insert(END, "\n")

        log_window['state'] = 'disabled'
        output['state'] = 'disabled'            

def restart_program():   

    log_window['state'] = 'normal'
    items_to_order_plans['state'] = 'normal'
    output['state'] = 'normal' 

    
    # resetovani stage programu do pocatecniho stavu.
    stage.clear()
    stage.append(1)

    # resetovani popisu start buttonu do pocatecniho stavu.
    run_btn_text.set("START")

    log_window.delete(1.0, END)
    items_to_order_plans.delete(1.0, END)
    output.delete(1.0, END)

    log_window['state'] = 'disabled'
    items_to_order_plans['state'] = 'disabled'
    output['state'] = 'disabled'       

def copy_items_order_plan():
    obsah = items_to_order_plans.get(1.0, END)
    items_to_order_plans.clipboard_clear()
    items_to_order_plans.clipboard_append(obsah)

def copy_vysledek():
    obsah = output.get(1.0, END)
    output.clipboard_clear()
    output.clipboard_append(obsah)


root = Tk()
root.title("Převody PZN105 <-> PZN100")
root.iconbitmap("Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\Prevody_PZN100_PZN105\\graphics\\icon1.ico")
root.geometry('1200x620+0+3')

### Postup programem.
# Krok 1.
stage = [1]

# Empty space labels. ( Na vyplneni prazdnych mist)
empty1 = Label(root, height=2, width=5, text="", pady=7)
empty2 = Label(root, height=4, width=5, text="", pady=6)
empty3 = Label(root, height=1, width=5, text="", pady=8)
empty4 = Label(root, height=2, width=10, text="", padx=3)
empty5 = Label(root, height=0, width=10, text="", padx=3)
empty6 = Label(root, height=0, width=1, text="", padx=0)

# Rovna se bunka.
equals = StringVar()
equals.set("=")
equal = Label(root, textvariable=equals, borderwidth=5)

# Okno zadavani poctu dalsich dnu k proverovani.
dnu_proverit = StringVar()
dnu_proverit.set(0)

next_cal_days_to_check = Entry(root, width = 5, textvariable=dnu_proverit, justify=CENTER, relief=SUNKEN, borderwidth = 5, font=('Calibry 11'))
next_cal_days_to_check_tooltip = StringVar()
next_cal_days_to_check_tooltip.set(f'Kal. dnů\nprověřit:')
next_cal_days_to_check_tooltip_label = Label(root, textvariable=next_cal_days_to_check_tooltip, font=('Calibry 8'), justify=LEFT)

# Okno zobrazujici datum, do kdy se bude proverovat na zaklade udaje vyse.
do_datumu = StringVar()
# do_datumu.set(datetime.date.today().strftime("%d/%m/%Y").replace("/", "."))

check_to_date = Label(root, width = 8, textvariable=do_datumu, height=1, font=('Calibry 11'))
check_to_date_tooltip = StringVar()
check_to_date_tooltip.set("Do datumu:")
check_to_date_tooltip_label = Label(root, width=12, height=2, textvariable=check_to_date_tooltip, justify=LEFT, font=('Calibry 8'))

# Empty 1
empty1.grid(row=0,column=0,rowspan=1, columnspan=1)

# Seznam polozek k provereni
items_to_order_plans = Text(root, width=16, height=14, borderwidth = 5, font=('Calibry 11'))
items_to_order_plans['state'] = 'disabled'
items_to_order_plans_tooltip = StringVar()
items_to_order_plans_tooltip.set(f'Seznam itemů zadat\n↓ do CQ order plánů: ↓')
items_to_order_plans_tooltip_label = Label(root, width=25, padx=0, textvariable=items_to_order_plans_tooltip, justify=LEFT, font=('Calibry 8'))

# Program log.
log_window = Text(root, width=125, height=8, borderwidth = 5, font=('Calibry 8'))
log_window['state'] = 'disabled'
log_window_tooltip = StringVar()
log_window_tooltip.set(f'↓ Info programu ↓')
log_window_tooltip_label = Label(root, height=2, textvariable=log_window_tooltip, pady=1, font=('Calibry 8'))

# Output window.
output = Text(root, width=120, height=22, borderwidth = 5, font=('Calibry 11'))
output['state'] = 'disabled'
output_tooltip = StringVar()
output_tooltip.set(f'↓ Output window ↓')
output_label = Label(root, textvariable=output_tooltip, width=23, height=2, justify=LEFT, font=('Calibry 8'))

### Tlacitka
# Spusteni.
run_btn_text = StringVar()
run_btn_text.set("SPUSTIT")
button_run = Button(root, textvariable=run_btn_text, height=3, width=10, padx=5, pady=5, borderwidth=5, command= lambda: run_program())

# Restart.
button_restart = Button(root, height=1,width=10, text= f'RESTART', padx=5, pady=10, borderwidth=5, command=restart_program)

# Zkopirovani do schranky itemy do Order planu.
button_items_order_plan = Button(root, width=16, text= f'↑ COPY ↑', padx=5, pady=5, borderwidth=5, command=copy_items_order_plan)

# Zkopirovani do schranky vysledek.
button_copy_output = Button(root, height=3,width=16, text= f'→ COPY →', padx=5, pady=5, borderwidth=5, command=copy_vysledek)



# ### Usporadani mrizky.
# 
# # Dni k provereni.
next_cal_days_to_check_tooltip_label.grid(row=0, column=0, rowspan=1)
next_cal_days_to_check.grid(row=1, column= 0, columnspan=1, padx=5, pady=5)
 
# Rovna se.
equal.grid(row= 1, column=1)
 
# Datum do kdy proverovat.
check_to_date_tooltip_label.grid(row=0, column=2, padx=5, pady=5)
check_to_date.grid(row=1, column=2, padx=5, pady=5)
# 
# # Empty.
empty1.grid(row=2,column=0, rowspan=1)
empty4.grid(row=2,column=2)

# Program log window.
log_window_tooltip_label.grid(row=0, column=6, columnspan=1)
log_window.grid(row=1, column=6, rowspan=3, columnspan=60, padx=0, pady=0)
# 
# Itemy vyjet z order planu.
items_to_order_plans_tooltip_label.grid(row=3, column=0, columnspan=3, padx=0)
items_to_order_plans.grid(row=4, column=0, rowspan=1, columnspan=3, padx=0, pady=0)
# empty5.grid(row=5, column=0, columnspan=3)
# 
# Tlacitka.
empty2.grid(row=0,column=3, rowspan=2)
button_run.grid(row=0, column=4, pady=5, rowspan=2)
button_restart.grid(row=2,column=4)
empty3.grid(row=0,column=5, rowspan=1)
button_items_order_plan.grid(row=5, column=0, columnspan=3, padx=10)
button_copy_output.grid(row=6, column=0, rowspan=2, columnspan=3, pady=10)
empty6.grid(row=9, column=0, rowspan=3)

# Output.
output_label.grid(row=3, column=3, columnspan=2)
output.grid(row=4, column=3,rowspan=4, columnspan=60, padx=5, pady=5)

root.mainloop()
