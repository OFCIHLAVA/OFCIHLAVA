import datetime
from webbrowser import get
from prevody_data.excel_data import dnesni_datum

def planned_available_na_skladu(shortage_linky, order_plan_skladu, zahlavi_vystupu): # Doplneni Planned available daneho itemu v dane PDD podle linky. (realny stav / shortage na PZN105:)
    for line in shortage_linky:
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
                if order_type.upper() == "PLANNED PURCHASE ORDER":
                    order_type = ""
                    continue
                datum = order_plan_skladu.get(vrchol).get(linka).get("Date")            
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

def next_planned_available_date_not_shortage_sklad(shortage_linky, order_plan_skladu, zahlavi_vystupu):# Doplneni nejblizsiho datumu + PA, kdy bude Planned Available na skladu pro dany item v lince alespon 0 nebo vyssi.

    for line in shortage_linky:
        vrchol = line[zahlavi_vystupu.index("Item")]
        if vrchol == "Item": # Neresit linku zahlavi.
            continue

        shortage_planned_available_skladu = float(line[zahlavi_vystupu.index("Planned available Qty on PZN105 at PDD")])
        if shortage_planned_available_skladu >= 0: # Neresit linky, kde neni shortage v PDD linky.
            line.append("-")
            continue
        
        pdd = line[zahlavi_vystupu.index("Planned Delivery Date")].split(".")
        pdd = datetime.date(int(pdd[-1]), int(pdd[1]), int(pdd[0]))     
        vrchol_available_qty_sklad = 0
        
        if order_plan_skladu.get(vrchol) != None:
            for linka in range(1,len(order_plan_skladu.get(vrchol))+1):
                order_type = order_plan_skladu.get(vrchol).get(linka).get("Order type txt")
                if order_type.upper() == "PLANNED PURCHASE ORDER":
                    order_type = ""
                    continue
                datum = order_plan_skladu.get(vrchol).get(linka).get("Date")
                balance = order_plan_skladu.get(vrchol).get(linka).get("Transaction type txt")
                quantity = float(order_plan_skladu.get(vrchol).get(linka).get("Quantity").replace(",",""))
                if balance == "+ (Planned Receipt)":
                    vrchol_available_qty_sklad += quantity
                elif balance == "- (Planned Issue)": 
                    vrchol_available_qty_sklad -= quantity
                else:
                    print('POZOR - ERROR v +/- balance na u itemu {vrchol} na lince {linka}.')            
                
                if order_type.upper() == "STOCK" and len(order_plan_skladu.get(vrchol)) == 1:
                    line.append(f'{dnesni_datum().strftime("%d/%m/%Y").replace("/",".")}, PA Qty: {vrchol_available_qty_sklad}')
                    break
                elif vrchol_available_qty_sklad >= 0 and datum > pdd:
                    line.append(f'{datum.strftime("%d/%m/%Y").replace("/",".")}, PA Qty: {vrchol_available_qty_sklad}')
                    break
            else:
                line.append(f'Neexistuje')              
        else:
            line.append(f'No data.')   

def next_planned_available_date_simulate_prevody(shortage_linky, order_plan_skladu_kam_chchi_prevadet, order_plan_skladu_odkud_chchi_prevadet, zahlavi_vystupu):

    ### Pomocne pocitadlo simulovanych prevodu. (Ceho, na kdy a kolik qty uz jsem jakoby prevedl)
    uz_simulovano_prevod = dict()

    for line in shortage_linky:
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
            #print(f'Nelze prevest, na druhem skladu neni dostatecne mnozstvi.\n')
            line.append("Nelze prevest, na druhem skladu neni dostatecne mnozstvi.")
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
                line.append(f'Nelze prevest, na druhem skladu item {vrchol} vubec neni.')
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
                    line.append(f'Nelze prevest! {vrchol} {sum_planned_available_kam_prevadim_opraveno_o_uz_simulovane} Qty na {pdd_linky}. Linky v op100 {shortage_linky_pri_prevodu} by se dostaly do minusu.')