import datetime
from webbrowser import get
from prevody_data.excel_data import dnesni_datum

# Obsolete analyza funkce:

def planned_available_na_skladu(vrchol, order_plan_vrcholu): # Doplneni Planned available daneho itemu v dane PDD podle linky. (realny stav / shortage na PZN105:)

    order_type = ""    
    vrchol_available_qty_na_skladu = 0       
    for linka in order_plan_vrcholu:
        order_type = order_plan_vrcholu.get(linka).get("Order type txt")
        if order_type.upper() == "PLANNED PURCHASE ORDER":
            continue
        balance = order_plan_vrcholu.get(linka).get("Transaction type txt")
        quantity = float(order_plan_vrcholu.get(linka).get("Quantity").replace(",",""))                
        if balance == "+ (Planned Receipt)":
            vrchol_available_qty_na_skladu += quantity
        elif balance == "- (Planned Issue)": 
            vrchol_available_qty_na_skladu -= quantity
        else:
            return f'POZOR - ERROR order planu v +/- balance na u itemu {vrchol} na lince {linka}.'
    return vrchol_available_qty_na_skladu

def realny_demand_na_skladu(vrchol, order_plan_vrcholu): # Spocitani kolik demand je na sklade pouze planned a bude se muset uspokojit nejakou zatim nezaplanovanou orderou.

    real_demand_order_types = ["PLANNED PURCHASE ORDER", "PLANNED DISTRIBUTION ORDER", "PLANNED PRODUCTION ORDER", "PURCHASE ORDER ADVICE"]
 
    realny_demnad_na_skladu = 0       
    for linka in order_plan_vrcholu:   
        order_type = order_plan_vrcholu.get(linka).get("Order type txt")
        # Pokud to neni real deman typ linky, preskocit na dalsi.
        if order_type.upper() not in real_demand_order_types:
            continue
        balance = order_plan_vrcholu.get(linka).get("Transaction type txt")
        # Pokud je to planned transakce demandova (-), tu taky preskocit, resime pouze supply stranu.
        if balance == "- (Planned Issue)":
            continue
        quantity = float(order_plan_vrcholu.get(linka).get("Quantity").replace(",",""))                
        if balance == "+ (Planned Receipt)":
            realny_demnad_na_skladu += quantity
        elif balance == "- (Planned Issue)": 
            realny_demnad_na_skladu -= quantity
        else:
            return f'POZOR - ERROR order planu v +/- balance na u itemu {vrchol} na lince {linka}.'
    return realny_demnad_na_skladu 

def po_in_process_skladu(vrchol, order_plan_vrcholu): # Spocitani sumy qty na already na ceste pro dany item.

    po_in_process_order_types = ["PURCHASE ADVICE", "PURCHASE ORDER"]
 
    po_in_process_na_skladu = 0       
    for linka in order_plan_vrcholu:   
        order_type = order_plan_vrcholu.get(linka).get("Order type txt")
        # Pokud to neni spravny order typ linky, preskocit na dalsi.
        if order_type.upper() not in po_in_process_order_types:
            continue
        balance = order_plan_vrcholu.get(linka).get("Transaction type txt")
        # Pokud je to planned transakce demandova (-), tu taky preskocit, resime pouze supply stranu.
        if balance == "- (Planned Issue)":
            continue
        quantity = float(order_plan_vrcholu.get(linka).get("Quantity").replace(",",""))                
        if balance == "+ (Planned Receipt)":
            po_in_process_na_skladu += quantity
        elif balance == "- (Planned Issue)": 
            po_in_process_na_skladu -= quantity
        else:
            return f'POZOR - ERROR order planu v +/- balance na u itemu {vrchol} na lince {linka}.'
    return po_in_process_na_skladu 

# Prevody funkce:

def planned_available_na_skladu_datum(shortage_linky, order_plan_skladu, zahlavi_vystupu): # Doplneni Planned available daneho itemu v dane PDD podle linky. (realny stav / shortage na PZN105:)
    for line in shortage_linky:
        # Vyrvori seznam linek order planu, kde jsou purchase ordery, ktere uz se pocitaji do planned available, ale jeste nedorazily k nam.
        todays_purchase_orders_counted_but_not_yet_here = list()
        order_type = ""    
        vrchol = line[zahlavi_vystupu.index("Item")]
        if vrchol == "Item":
            continue
        pdd = line[zahlavi_vystupu.index("Planned Delivery Date")].split(".")
        pdd = datetime.date(int(pdd[-1]), int(pdd[1]), int(pdd[0]))
        vrchol_available_qty_na_skladu = 0
        if order_plan_skladu.get(vrchol) != None:
            for linka in order_plan_skladu.get(vrchol):
                order_type = order_plan_skladu.get(vrchol).get(linka).get("Order type txt")
                datum = order_plan_skladu.get(vrchol).get(linka).get("Date")

                # if order_type.upper() == "PURCHASE ORDER":
                    # print(vrchol, datum, dnesni_datum())
                if order_type.upper() == "PLANNED PURCHASE ORDER":
                    order_type = ""
                    continue
                elif order_type.upper() == "PURCHASE ORDER" and datum == dnesni_datum():
                    todays_purchase_orders_counted_but_not_yet_here.append(f'{order_plan_skladu.get(vrchol).get(linka).get("Order Number")}, Qty {round(float(order_plan_skladu.get(vrchol).get(linka).get("Quantity")), 1)}')
                    # print(todays_purchase_orders_counted_but_not_yet_here)
                if datum <= pdd:
                    balance = order_plan_skladu.get(vrchol).get(linka).get("Transaction type txt")
                    quantity = float(order_plan_skladu.get(vrchol).get(linka).get("Quantity").replace(",",""))                
                    if balance == "+ (Planned Receipt)":
                        vrchol_available_qty_na_skladu += quantity
                    elif balance == "- (Planned Issue)": 
                        vrchol_available_qty_na_skladu -= quantity
                    else:
                        print('POZOR - ERROR order planu v +/- balance na u itemu {vrchol} na lince {linka}.')
        line.append(float(vrchol_available_qty_na_skladu))
        if len(todays_purchase_orders_counted_but_not_yet_here) != 0:
            line.append("; ".join(todays_purchase_orders_counted_but_not_yet_here))
        else:
            line.append("-")

def next_planned_available_date_not_shortage_sklad(shortage_linky, order_plan_skladu, zahlavi_vystupu):# Doplneni nejblizsiho datumu + PA, kdy bude Planned Available na skladu pro dany item v lince alespon 0 nebo vyssi.

    for line in shortage_linky:
        # print(f'PRED next planed not shortage\n {str(order_plan_skladu)}\n line:\n {line})')
        vrchol = line[zahlavi_vystupu.index("Item")]
        if vrchol == "Item": # Neresit linku zahlavi.
            continue

        shortage_planned_available_skladu = float(line[zahlavi_vystupu.index("Planned available Qty on PZN105 at PDD")])
        if shortage_planned_available_skladu >= 0: # Neresit linky, kde neni shortage v PDD linky.
            line.append("-")
            continue
        
        pdd = line[zahlavi_vystupu.index("Planned Delivery Date")].split(".")
        pdd = datetime.date(int(pdd[-1]), int(pdd[1]), int(pdd[0]))     
        
        if order_plan_skladu.get(vrchol) != None:            
            vrchol_available_qty_sklad = 0
            for linka in range(1,len(order_plan_skladu.get(vrchol))+1):
                
                order_type = order_plan_skladu.get(vrchol).get(linka).get("Order type txt")
                if order_type.upper() == "PLANNED PURCHASE ORDER":
                    if len(order_plan_skladu.get(vrchol)) > 1:
                        order_type = ""
                        continue
                    else:
                        line.append(f'Neexistuje')
                        break
                datum = order_plan_skladu.get(vrchol).get(linka).get("Date")
                balance = order_plan_skladu.get(vrchol).get(linka).get("Transaction type txt")
                quantity = float(order_plan_skladu.get(vrchol).get(linka).get("Quantity").replace(",",""))
                if balance == "+ (Planned Receipt)":
                    vrchol_available_qty_sklad += quantity
                elif balance == "- (Planned Issue)": 
                    vrchol_available_qty_sklad -= quantity
                else:
                    print('POZOR - ERROR v +/- balance na u itemu {vrchol} na lince {linka}.')            
                # print(vrchol, linka, len(order_plan_skladu.get(vrchol)), vrchol_available_qty_sklad)
                # Moznosti vysledku:
                # 1. Pokud pouze stock → ok.
                if order_type.upper() == "STOCK" and len(order_plan_skladu.get(vrchol)) == 1:
                   line.append(f'{dnesni_datum().strftime("%d/%m/%Y").replace("/",".")}, PA Qty: {vrchol_available_qty_sklad}')
                   break
                # 2. Pokud po projiti linky je Planned available >= 0 a datum je vetsi nez PDD, → tak datum (neni treba resit pro PDD, dela se jen pro linky, kde uz vim, ze PA v PDD je < 0).
                elif vrchol_available_qty_sklad >= 0 and datum > pdd:                    
                    line.append(f'{datum.strftime("%d/%m/%Y").replace("/",".")}, PA Qty: {vrchol_available_qty_sklad}')
                    break                
                # 3. Pokud jsem prosel vsechny linky a nespustila se zadna z podminek vyse.
                elif linka == len(order_plan_skladu.get(vrchol)):
                    # print(f'TUDY')
                    # Pokud Planned available na posledni lince >= 0.
                    if vrchol_available_qty_sklad >= 0:
                        # a) Pokud datum posledni linky je pred PDD, ale je az po dnesnim datu → datum posledni linky.
                        if datum < pdd and datum > dnesni_datum():
                            line.append(f'{datum.strftime("%d/%m/%Y").replace("/",".")}, PA Qty: {vrchol_available_qty_sklad}')
                            break
                        # b) Pokud datum posledni linky je pred PDD a mensi rovno dnesni datum → dnesni datum.
                        elif datum < pdd and datum <= dnesni_datum():
                            line.append(f'{dnesni_datum().strftime("%d/%m/%Y").replace("/",".")}, PA Qty: {vrchol_available_qty_sklad}')
                            break                            
                        # c) Pokud datum posledni linky je po PDD → datum linky.
                        else:                        
                            line.append(f'{datum.strftime("%d/%m/%Y").replace("/",".")}, PA Qty: {vrchol_available_qty_sklad}')
                            break
            else:
                line.append(f'Neexistuje')

        else:
            line.append(f'No data.')   
        # print(f'PO next planed not shortage\n {str(order_plan_skladu)}\n line:\n {line})')

def next_planned_available_date_simulate_prevody(shortage_linky, order_plan_skladu_kam_chchi_prevadet, order_plan_skladu_odkud_chchi_prevadet, zahlavi_vystupu):

    ### Pomocne pocitadlo simulovanych prevodu. (Ceho, na kdy a kolik qty uz jsem jakoby prevedl)
    uz_simulovano_prevod = dict()

    for line in shortage_linky:
        # print(f'LINKA PRED mozno prevest?:\n {line}')
        vrchol = line[zahlavi_vystupu.index("Item")]
        if vrchol == "Item":
            continue  
        
        pdd_linky = line[zahlavi_vystupu.index("Planned Delivery Date")].split(".")
        pdd_linky = datetime.date(int(pdd_linky[-1]), int(pdd_linky[1]), int(pdd_linky[0]))

        ### Sum Qty uz nasimulovanych prevodu daneho vrcholu s datumy <= PDD linky.            
        uz_simulovano_qty_vrcholu = float()
            ### Pokud uz se u vrcholu simulovaly nejake provody:
        if uz_simulovano_prevod.get(vrchol):
            # print(f'uz simulovano {uz_simulovano_prevod.get(vrchol)}')
            ### Projdi vsechny datumy mensi rovno PDD linky a nacti jejich Sumu.
            for prevod in uz_simulovano_prevod.get(vrchol):
                # print(f'prevod uz {prevod}')
                if prevod <= pdd_linky:
                    uz_simulovano_qty_vrcholu += float(uz_simulovano_prevod.get(vrchol).get(prevod))
                    # print(f'Qty uz {uz_simulovano_prevod.get(vrchol).get(prevod)}')
                    # print(f'uz simulovano v pdd {uz_simulovano_qty_vrcholu}')
        ### Kolik chybi Planned available na dane lince.
        sum_planned_available_kam_prevadim = float(line[zahlavi_vystupu.index("Planned available Qty on PZN105 at PDD")])
        ### Opravene o to, kolik daneho itemu jsem uz prevedl.
        sum_planned_available_kam_prevadim_opraveno_o_uz_simulovane = sum_planned_available_kam_prevadim + uz_simulovano_qty_vrcholu
        #print(f'Opravene PA 105: {sum_planned_available_kam_prevadim_opraveno_o_uz_simulovane}')

        ### 1. Overeni, zda je planned available linky, kterou kontroluji, opravene o uz prevedene qty < 0.
        if  sum_planned_available_kam_prevadim_opraveno_o_uz_simulovane >= 0: ### 1A) Pokud neni planned available mensi 0 → neni treba nic prevadet. Preskoci na dalsi linku.
            # print(f'OK - Neresit.')
            line.append("-")
            continue
        ### 1B) Pokud je planned available mensi 0 → Proveri se, zda je mozno sem prevest z druheho skladu opravene o mnozstvi, ktere jsem uz prevedl.
        sum_planned_available_odkud_prevadim = float(line[zahlavi_vystupu.index("Planned available Qty on PZN100 at PDD")]) - uz_simulovano_qty_vrcholu
        # print(f'PA na 100 na PDD: {sum_planned_available_odkud_prevadim}')
        ### 2. Podivej se, jestli je alespon tolik, kolik bych potreboval prevest k dispozici na sklade, odkud prevadim v PDD.
        if sum_planned_available_kam_prevadim + sum_planned_available_odkud_prevadim < 0: ### 2a] Pokud neni, nelze prevadet. → preskocit na dalsi linku.
            # print(f'GUGU')
            #print(f'Nelze prevest, na druhem skladu neni dostatecne mnozstvi.\n')
            line.append("Nelze, na druhem skladu neni dostatecne mnozstvi.")
            continue        
        else: ### 2b] Pokud je, → nasimulovat prevod.
            # print(f'PA na 100 je dost → koukam se, co ostatni linky na pzn100.')          
            simulovane_planned_available_qty_sklad_odkud_prevadim = 0
            shortage_linky_pri_prevodu = list()           

            ### ziskani Supply time vrcholu.
            if order_plan_skladu_kam_chchi_prevadet.get(vrchol):
                supply_time_vrcholu = float(order_plan_skladu_kam_chchi_prevadet.get(vrchol).get(1).get("Supply LT [work days]"))
                ### prevedeni na cal days
                supply_time_vrcholu = supply_time_vrcholu/5*7
            elif order_plan_skladu_odkud_chchi_prevadet.get(vrchol):
                supply_time_vrcholu = float(order_plan_skladu_odkud_chchi_prevadet.get(vrchol).get(1).get("Supply LT [work days]"))
                ### prevedeni na cal days
                supply_time_vrcholu = supply_time_vrcholu/5*7
            else:
                supply_time_vrcholu = "N/A"        
            # print(f'Supply time vrcholu cal days: {supply_time_vrcholu}.')
            ### Pokud neni mozne zjistit z dat order planu supply LT → linka nelze resit. Preskoci na dalsi linku.    
            if supply_time_vrcholu == "N/A":
                line.append(f'Nelze urcit - neznamy suppply time vrcholu {vrchol}')
                continue
            
            ### Jinak prover linky order planu odkud chci prevadet s datumem mensim nez dnes + doba za jakou to nakup nakoupi + safety time 5 dni, jestli se prevodem nedostanou do minusu.  
            datum_proverovat_do = dnesni_datum() + datetime.timedelta(supply_time_vrcholu) + datetime.timedelta(5)
            # print(f'datum linek pzn 100 proverovat do {datum_proverovat_do}.')
            ### Pokud v order planu okdud chci prevadet neni zadna linka itemu → nelze prevadet. Preskoci na dalsi linku.
            if not order_plan_skladu_odkud_chchi_prevadet.get(vrchol):                
                line.append(f'Nelze, na druhem skladu item {vrchol} vubec neni.')
                continue
            ### Jinak simuluj prevod.
            else:
                # print(f'Simuluji prevod {abs(sum_planned_available_kam_prevadim_opraveno_o_uz_simulovane)} Qty z {100} na {105}...')
                ### Postupne projdi vsechny linky od 1 (nejstarsi) az po posledni (nejmladsi).
                for linka in range(1,len(order_plan_skladu_odkud_chchi_prevadet.get(vrchol))+1):
                    # print(f'\nkoukam na {linka}. linku {vrchol} na pzn 100...')
                    datum = order_plan_skladu_odkud_chchi_prevadet.get(vrchol).get(linka).get("Date")
                    order_number = order_plan_skladu_odkud_chchi_prevadet.get(vrchol).get(linka).get("Order Number")
                    # print(f'Order number: {order_number}.')
                    # print(f'datum linky op: {datum}.')
                    # print(f'datum linek pzn 100 proverovat do {datum_proverovat_do}.')
                    if datum < datum_proverovat_do:
                        # print("ANO datum linky je mensi nez datum do kdy proverovat.") 
                        order_type = order_plan_skladu_odkud_chchi_prevadet.get(vrchol).get(linka).get("Order type txt")
                        # print(f'Order type: {order_type}.')
                        ### Planned purchase orders se nepocitaji. → preskocit na dalsi linku order planu.
                        if order_type.upper() == "PLANNED PURCHASE ORDER":
                            order_type = ""
                            print(f'HOHO')
                            continue                   
                        balance = order_plan_skladu_odkud_chchi_prevadet.get(vrchol).get(linka).get("Transaction type txt")
                        # print(f'pohyb: {balance}.')
                        quantity = float(order_plan_skladu_odkud_chchi_prevadet.get(vrchol).get(linka).get("Quantity").replace(",",""))
                        # print(f'QTY: {quantity}.')
                        # print(quantity, type(quantity))
                        if balance == "+ (Planned Receipt)":
                            simulovane_planned_available_qty_sklad_odkud_prevadim += quantity
                        elif balance == "- (Planned Issue)": 
                            simulovane_planned_available_qty_sklad_odkud_prevadim -= quantity
                        else:
                            print('POZOR - ERROR v +/- balance na u itemu {vrchol} na lince {linka}.')
                        # print(f'kontrola available qty v PDD bez opravy sim. prevodu: {simulovane_planned_available_qty_sklad_odkud_prevadim}.')
                        # print(f'kontrola available qty v PDD vcetne opravy sim. prevodu: {simulovane_planned_available_qty_sklad_odkud_prevadim + sum_planned_available_kam_prevadim_opraveno_o_uz_simulovane}.')
                        # print(f'planned_available_pzn105 {planned_available_pzn105}'         
                        ### Pro kazdou linku order planu, okdud chci prevadet zkontrolovat, jestli se prevodem pro linky s datumem < dnes + supply time + safety time nedostanu do minusu.
                        if simulovane_planned_available_qty_sklad_odkud_prevadim + sum_planned_available_kam_prevadim_opraveno_o_uz_simulovane < 0:                        
                                # print(f'Shortage by byl na lince order planu {linka}, {simulovane_planned_available_qty_sklad_odkud_prevadim + sum_planned_available_kam_prevadim_opraveno_o_uz_simulovane} Qty.')
                                shortage_linky_pri_prevodu.append((datum, order_number, simulovane_planned_available_qty_sklad_odkud_prevadim + sum_planned_available_kam_prevadim_opraveno_o_uz_simulovane))
                        # else:
                            # print(f'Linka {linka} op OK. Po prevodu by bylo PA {simulovane_planned_available_qty_sklad_odkud_prevadim + sum_planned_available_kam_prevadim_opraveno_o_uz_simulovane}')    
                    # else:
                        # print(f'Linka ma datum dosta daleko → Neresime.')
                # print(f'Shortage linky pri simulaci prevodu {vrchol} {shortage_linky_pri_prevodu, len(shortage_linky_pri_prevodu)}')    
                if len(shortage_linky_pri_prevodu) == 0:
                    # print(f'Mozno prevest! {vrchol} {sum_planned_available_kam_prevadim_opraveno_o_uz_simulovane} Qty na {pdd_linky}')
                    # print(uz_simulovano_prevod)
                        
                    ### Pridani simulovaneho mnozstvi k ostatnim prevodum.
                    ### prevod stejneho vrcholu se jeste nesimuloval.
                    if not uz_simulovano_prevod.get(vrchol):
                        # print(f'{vrchol} se jeste neprevadel. Pridavam {abs(sum_planned_available_kam_prevadim_opraveno_o_uz_simulovane)} Qty do uz simulovanych prevodu s datumem {pdd_linky}.')
                        pdd_dict = dict()
                        pdd_dict[pdd_linky] = abs(sum_planned_available_kam_prevadim_opraveno_o_uz_simulovane)
                        uz_simulovano_prevod[vrchol] = pdd_dict
                        # print(uz_simulovano_prevod)  
                    ### prevod stejneho vrcholu uz se simuloval.
                    else:
                        # print(f'{vrchol} se uz prevadel. Pridavam {abs(sum_planned_available_kam_prevadim_opraveno_o_uz_simulovane)} Qty do uz simulovanych prevodu s datumem {pdd_linky}.')
                        ### jeste ne neprevadel v PDD linky.    
                        if not uz_simulovano_prevod.get(vrchol).get(pdd_linky):
                            # print(f'V datum {pdd_linky} se jeste neprevadel. Pridavam {abs(sum_planned_available_kam_prevadim_opraveno_o_uz_simulovane)} Qty do uz simulovanych prevodu s datumem {pdd_linky}.')
                            uz_simulovano_prevod[vrchol][pdd_linky] = abs(sum_planned_available_kam_prevadim_opraveno_o_uz_simulovane)
                            # print(uz_simulovano_prevod)
                        ### uz se nejake mnozstvi v tento den prevadelo.  
                        else:
                            # print(f'V datum {pdd_linky} se uz neco prevadelo. Scitam {abs(sum_planned_available_kam_prevadim_opraveno_o_uz_simulovane)} Qty do uz s timto datumem {pdd_linky}.')
                            uz_simulovano_prevod[vrchol][pdd_linky] = uz_simulovano_prevod.get(vrchol).get(pdd_linky) + abs(sum_planned_available_kam_prevadim_opraveno_o_uz_simulovane)
                            # print(uz_simulovano_prevod)
                    ### Pripojeni vyslednu na konec linky.
                    line.append(f'Prevest {abs(sum_planned_available_kam_prevadim_opraveno_o_uz_simulovane)} z PZN100.')
                else:
                    # print(f'Nelze prevest! {vrchol} {sum_planned_available_kam_prevadim_opraveno_o_uz_simulovane} Qty na {pdd_linky}. Linky v op100 {shortage_linky_pri_prevodu} by se dostaly do minusu.')
                    ### Pripojeni vyslednu na konec linky.
                    line.append(f'Nelze! {vrchol} {abs(sum_planned_available_kam_prevadim_opraveno_o_uz_simulovane)} Qty na {pdd_linky}. Linky v op100 {shortage_linky_pri_prevodu} by se dostaly do minusu.')
        # print(f'LINKA PO mozno prevest?:\n {line}')

def location_to_prevest_from(shortage_linky, warehouse_locations, zahlavi_vystupu): # Doplneni sloupce s informaci, z jake lokace nechat dil prevest.
    # print(zahlavi_vystupu)
    # 
    # print(f'2. KROK\n\n')

    for line in shortage_linky:
        # print(f'Linka PRED: {line}')
        # print(f'len line {len(line)}')
        # for zahlavi, udaj in enumerate(line):
            # print(zahlavi_vystupu[zahlavi], udaj,sep=":")
        # Prevadene linky pozname podle textu "Prevest" v sloupci Mozno prevest. 
        # print(zahlavi_vystupu.index("Mozno prevest z PZN100 aniz by se ohrozily budouci linky na PZN100?"))
        mozno_prevest = line[zahlavi_vystupu.index("Mozno prevest z PZN100 aniz by se ohrozily budouci linky na PZN100?")]
        prevest = False
        if "Prevest" in mozno_prevest:
            prevest = True
        if prevest:
            vrchol = line[zahlavi_vystupu.index("Item")]
            # Projit warehouse locations pro dany item a najit lokaci s nejvetsi IOH po zohledneni uz prevadenych Qty z dane lokace.
            max_ioh = 0
            max_ioh_location = ""

            # test zda je vrchol v lokacich na skladu
            if warehouse_locations.get(vrchol) == None:
                line.append("POZOR! rika to ze prevod, ale lokace na skladu jso prazdne!")      
            else:
                # print(warehouse_locations.get(vrchol))
                for linka in warehouse_locations.get(vrchol):
                    linka_ioh = float(warehouse_locations.get(vrchol).get(linka).get("Inventory on hand").replace(",",""))
                    linka_location = warehouse_locations.get(vrchol).get(linka).get("Location")

                    if linka_ioh > max_ioh:
                        max_ioh = linka_ioh
                        max_ioh_location = linka_location
                line.append(max_ioh_location)
        else:
            line.append("-")
        # print(f'Linka PO: {line}')        