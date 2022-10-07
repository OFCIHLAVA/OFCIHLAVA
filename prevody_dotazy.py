import datetime
from webbrowser import get

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