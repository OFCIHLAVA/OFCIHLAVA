import openpyxl as excel

def je_to_man_placard(item, parameters):
    nazev = parameters.get(item).get("description")
    man_placard = "placar"
    instalace   = ["installation", "inst"] 
    
    test_man_placard = False
    
    if (man_placard.upper() in nazev.upper()) and (parameters.get(item).get("Item type") == "Manufactured"):
        test_man_placard = True
        for podminka in instalace:
            if podminka.upper() in nazev.upper():
                test_man_placard = False
                break
    return test_man_placard

def je_to_id_placard(item, parameters):
    podminka = "ID PLACARD"
    if podminka == parameters.get(item).get("description").strip():
        return True
    return False

def lokalni_dodavatele(item, parameters):
    lok_dodavatele_name = (
    "Carry Goods s.r.o.",
    "Elakov Production s.r.o.",
    "EVEKTOR-AEROTECHNIK a.s.",
    "HT Metal s.r.o.",
    "LABARA s.r.o.",
    "PRIMAPOL-METAL-SPOT, s.r.o.",
    "STARTECH s.r.o.",
    "Tuma aerospace s.r.o.",
    "Workpress Aviation s.r.o.",
    "Safran Cabin Lamphun Ltd.",
    "Molins a.s.",
    "SW-MOTECH s.r.o.",
    "STARTECH s.r.o.",
    "AXA CNC stroje, s.r.o."
    )

    if parameters.get(item).get("Dodavatel").strip() in lok_dodavatele_name:
        return True
    else:
        return False

# def zahlavi_data_txt(zahlavi): # Ziskani nazvu sloupcu jako list.  
#     zahlavi = []
#     for c in range(1,min(excel_sheet.max_column+1, 16384)):
#         column_name = excel_sheet.cell(8, c).value
#         # if column_name not in zahlavi and column_name:
#         zahlavi.append(column_name)
#     return zahlavi

# 1. Naceteni dat.

with open("BFE file_txt.txt", "r", encoding='utf-8') as file:
    data_import = file.readlines()

# 2. Rozdeleni dat na jednotliva pole podle "/t".

data = list()
errors = list()
for line in data_import:
    line = line.replace("\n","")
    line = line.split("\t")
    # print(len(line), line, sep=" ")
    data.append(line)

print(f'Data loaded ...')

# 3. Vytvoreni zahlavi.

zahlavi = data[0]
print(f'Zahlavi: {zahlavi}')

# 4. Samotna data bez zahlavi.

data_without_zahlavi = data[1:]

# 5. Kontrola mezer v Part numberu.

blanks_pn = set()

index_pn = zahlavi.index("Part number")
for line in data_without_zahlavi:
    if " " in line[index_pn]:
        line[index_pn] = line[index_pn].replace(" ","")
        blanks_pn.add(line[index_pn])
if len(blanks_pn) != 0:
    print(f'Nalezeno {len(blanks_pn)} P/N s mezerou v cisle:')
    for pn in blanks_pn:
        print('{pn}\n')


# 6. Pochytani linek dat, ktere neodpovidaji poctem poli dat poctu nazvu sloupcu v zahlavi. 

for line in data_without_zahlavi:
    if len(line) != len(zahlavi):
        errors.append(line)
        print(f'POZOR! Line {line} nesouhlasi pocet poli dat se zahlavim.')

# 6.1 Odstraneni linek s CH - costovymi polozkami.
print(f'Pocet linek v datech {len(data_without_zahlavi)}')
cost_lines_to_delete = list()

for line in data_without_zahlavi:
    item = line[zahlavi.index("Part number")]
    item_type = line[zahlavi.index("Item type")]
    
    if item[0:2] == "CH" and item_type == "Cost":
        cost_lines_to_delete.append(line)

print(f'Nalezeno a smazano {len(cost_lines_to_delete)} cost polozek:')
for line in cost_lines_to_delete:
    print(line[zahlavi.index("Part number")])
    data_without_zahlavi.remove(line)

print(f'Pocet linek v datech {len(data_without_zahlavi)}')

# 6.2 Odstraneni linek s "USE" polozkami.

print(f'Pocet linek v datech {len(data_without_zahlavi)}')
use_lines_to_delete = list()

for line in data_without_zahlavi:
    item_description = line[zahlavi.index("description")]
    
    if item_description[0:4].upper() == "USE ":
        use_lines_to_delete.append(line)

print(f'Nalezeno a smazano {len(use_lines_to_delete)} "USE" polozek:')
for line in use_lines_to_delete:
    print(line[zahlavi.index("Part number")])
    data_without_zahlavi.remove(line)

print(f'Pocet linek v datech {len(data_without_zahlavi)}')

# 7. Tisk vyslednych dat.

# for line in data_without_zahlavi:
#     print(line)

print("\nERRORS:")
if len(errors) >0:
    for line in errors:
        print(len(line), line, sep=" ")
else:
    print("OK - vsechny linky dat maji spravnou delku.\n")

# 8. Priprava databaze.

databaze_itemu = dict()

for line in data_without_zahlavi:
    item = str(line[zahlavi.index("Part number")])
    # print(f'Item: {item}')

    data_itemu = dict()
    for zahlavi_sloupce in zahlavi:
        data_itemu[zahlavi_sloupce] = line[zahlavi.index(zahlavi_sloupce)]
    # print(f'data itemu: {data_itemu}')
    databaze_itemu[item] = data_itemu

# 10. Kontrola spravnosti dat pro vsechny polozky.

# 10.0.5 Pridani informace, zda se jedna o MAN PLACARD / ID PLACARD.
zahlavi.append("MAN PLACARD / ID PLACARD?")
for item in databaze_itemu:
    pn_in_mam_placard = False
    pn_is_id_placard = False

    if je_to_man_placard(item, databaze_itemu):
        if je_to_id_placard(item, databaze_itemu):
            pn_is_id_placard = True
            databaze_itemu[item]["MAN PLACARD / ID PLACARD?"] = "ID PLACARD"
        else:
            pn_in_mam_placard = True
            databaze_itemu[item]["MAN PLACARD / ID PLACARD?"] = "MAN PLACARD"
    else:
        databaze_itemu[item]["MAN PLACARD / ID PLACARD?"] = "no"

# 10.1 Sales group.

zahlavi.append("Sales group to be")

for item in databaze_itemu:
    sales_group_to_be = "?"
    # print(item)
    # print(databaze_itemu.get(item).get("Sales group LN").upper())
    # Ponechat skupinu stejnou pro BLK, HARM a AFDAL polozky.    
    if "BLK" in databaze_itemu.get(item).get("Sales group LN").upper():
        sales_group_to_be = "BLK"

    elif "HARM" in databaze_itemu.get(item).get("Sales group LN").upper():
        sales_group_to_be = "HARM"
    elif "AFDAL" in databaze_itemu.get(item).get("Sales group LN").upper():
        sales_group_to_be = "AFDAL"
    
    # 26.09.2022 - Dle domluvy s PS a JV, nechat vsechny MK pro 2023 jako POR a dat jim "divnou SG - 001"
    elif "MK" in databaze_itemu.get(item).get("Sales group LN").upper():
        sales_group_to_be = "001"    
    # MK, 1, #N/A a N/A prepsat na spravnou skupinu podle Python, pripadne smazat na N/A.
    elif databaze_itemu.get(item).get("Sales group LN").upper() == "1" or "N/A" in databaze_itemu.get(item).get("Sales group LN").upper():
        if "N/A" in databaze_itemu.get(item).get("Sales group Python").upper():
            sales_group_to_be = "N/A"
        else:
            sales_group_to_be = databaze_itemu.get(item).get("Sales group Python")
    # BFE itemy prepsat na spravnou skupinu podle Python, pripadne nechat to co tam bylo.
    elif "BFE" in databaze_itemu.get(item).get("Sales group LN").upper():
        if "N/A" in databaze_itemu.get(item).get("Sales group Python").upper():
            sales_group_to_be = databaze_itemu.get(item).get("Sales group LN")
        else:
            sales_group_to_be = databaze_itemu.get(item).get("Sales group Python")
    # SFE itemy prepsat na spravnou skupinu podle Python, pripadne nechat to co tam bylo.
    elif "SFE" in databaze_itemu.get(item).get("Sales group LN").upper():
        if "N/A" in databaze_itemu.get(item).get("Sales group Python").upper():
            sales_group_to_be = databaze_itemu.get(item).get("Sales group LN")
        else:
            sales_group_to_be = databaze_itemu.get(item).get("Sales group Python")
    # MIX itemy prepsat na spravnou skupinu podle Python, pripadne nechat to co tam bylo.
    elif "MIX" in databaze_itemu.get(item).get("Sales group LN").upper():
        if "N/A" in databaze_itemu.get(item).get("Sales group Python").upper():
            sales_group_to_be = databaze_itemu.get(item).get("Sales group LN")
        else:
            sales_group_to_be = databaze_itemu.get(item).get("Sales group Python")

    databaze_itemu[item]["Sales group to be"] = sales_group_to_be

# 10.1.1 Upresneni zda se jedna o 01/02 skupinu.

for item in databaze_itemu:
    # Proverime linky kde to ma byt SFE/BFE.
    if "BFE" in databaze_itemu.get(item).get("Sales group to be").upper() or "SFE" in databaze_itemu.get(item).get("Sales group to be").upper():
        # Pro Purchased dily.
        if "PURCHASED" in databaze_itemu.get(item).get("Item type").upper():
            # Pokud neni "Planner" v nazvu nakupciho → 01.
            if "PLANNER" not in databaze_itemu.get(item).get("Nakupci").upper():            
            # Pokud se nejedna o polozku od lokalniho dodavatele → 01.    
                if not lokalni_dodavatele(item, databaze_itemu):
                    if "BFE" in databaze_itemu.get(item).get("Sales group to be").upper():
                        databaze_itemu[item]["Sales group to be"] = "BFE01"
                    else:
                        databaze_itemu[item]["Sales group to be"] = "SFE01"
                # Pokud se jedna o polozku od lokalniho dodavatele → 02.
                else:
                    if "BFE" in databaze_itemu.get(item).get("Sales group to be").upper():
                        databaze_itemu[item]["Sales group to be"] = "BFE02"
                    else:
                        databaze_itemu[item]["Sales group to be"] = "SFE02"
            # Pokud se jedna o Pruchased s "Planner" v nazvu nakupciho → 02.
            else:
                if "BFE" in databaze_itemu.get(item).get("Sales group to be").upper():
                    databaze_itemu[item]["Sales group to be"] = "BFE02"
                else:
                    databaze_itemu[item]["Sales group to be"] = "SFE02"
        # Pokud se jedna o "Manufactured" polozku → 02.
        else:
            if "BFE" in databaze_itemu.get(item).get("Sales group to be").upper():
                databaze_itemu[item]["Sales group to be"] = "BFE02"
            else:
                databaze_itemu[item]["Sales group to be"] = "SFE02"

# 10.2 Sales LT.
zahlavi.append("Sales LT to be")

for item in databaze_itemu:
    sales_lt_ln = databaze_itemu.get(item).get("LN Sales LT [days]")
    sales_lt_python = databaze_itemu.get(item).get("LN Sales LT [days] Python")

    # 14 cal dni pro MAN placardy a ID PLACARDY
    
    # rucne doplnene placardy, ktere nechytne funkce na PLACARDY.
    if item == "668273-103":
        sales_lt_to_be = str(14)
    elif je_to_man_placard(item, databaze_itemu):
        sales_lt_to_be = str(14)
    elif je_to_id_placard(item, databaze_itemu):
        sales_lt_to_be = str(14)
    # Pokud neni udaj z Pythonu → dat co je ted v LNku.
    elif sales_lt_python == "N/A":
        sales_lt_to_be = sales_lt_ln
    # Pokud je sales LN v LN pod 14 dni → dat co je v pythonu.
    elif float(sales_lt_ln) <14:
        sales_lt_to_be = round(float(sales_lt_python.replace(",",".")),0)
    # U vsech ostatnich porovnat lt v LN s Pythonem a dat z Pythonu pokud je vyssi nez z LN, jinak nechat LN LT.
    else:
        if float(sales_lt_ln.replace(",",".")) < float(sales_lt_python.replace(",",".")):
            sales_lt_to_be = round(float(sales_lt_python.replace(",",".")),0)
        else:
            sales_lt_to_be = sales_lt_ln

    databaze_itemu[item]["Sales LT to be"] = str(sales_lt_to_be).replace(".",",")

# 10.2.5 MOQ to be.
zahlavi.append("Sales MOQ to be")

for item in databaze_itemu:

    moq_to_be = 'N/A'

    # 1. MAN PLACARDY = 3, ID placardy = 1
    man_placard = je_to_man_placard(item, databaze_itemu)
    id_placard = je_to_id_placard(item, databaze_itemu)
    
    if man_placard:
        if id_placard:
            moq_to_be = 1
        else:
            moq_to_be = 3
        databaze_itemu[item]["Sales MOQ to be"] = str(moq_to_be)
    else:
        # print(f'neplacard')
        # print(databaze_itemu.get(item).get("MOQ LN"))
        # 2. Pokud neni MOQ neplatne (N/A).
        if "N/A" not in databaze_itemu.get(item).get("MOQ LN"):
            moq_now = float(databaze_itemu.get(item).get("MOQ LN").replace(",","."))
            # Pro MOQ 0 → dat 1.
            if moq_now == 0:
                moq_to_be = 1        
            elif moq_now % 1 != 0:
                moq_to_be = "POZOR, divne MOQ - proverit s Petrem, jestli dat MOQ 1"
            else:
                moq_to_be = 1
            databaze_itemu[item]["Sales MOQ to be"] = str(moq_to_be)
    
        else:
            databaze_itemu[item]["Sales MOQ to be"] = 'N/A'

# 10.3 Sales price actual now.
zahlavi.append("Actual Sales price now [EUR]")

for item in databaze_itemu:
    historical_eur_price = databaze_itemu.get(item).get("BFE historical 2022 [EUR]")
    pr_nac_eur_price = databaze_itemu.get(item).get("BFE pr. Nac 2022")

    # Pokud je platna cena z prubezneho nacenovani → cena z pr nac.
    if pr_nac_eur_price != "0" and "N/A" not in pr_nac_eur_price.upper():
        databaze_itemu[item]["Actual Sales price now [EUR]"] = pr_nac_eur_price
    else:
        # Pokud je platna cena z historicalu → cena z historicalu.
        if historical_eur_price != "0" and "N/A" not in historical_eur_price.upper() and "POR" not in historical_eur_price.upper():
            databaze_itemu[item]["Actual Sales price now [EUR]"] = historical_eur_price
        # Pokud ani jedno z vysse neplati → "N/A"
        else:
            databaze_itemu[item]["Actual Sales price now [EUR]"] = "N/A"

# 10.4 Marze.

zahlavi.append("Actual margin type")
zahlavi.append("Actual margin")
zahlavi.append("Low margin")

nizka_marze_hranice = 0.66

# 10.4.1 Purchase marze.

neplatni_dodavatele = ["0", "N/A", "SAFRAN CABIN LAMPHUN LTD."]
neplatna_cena = ["0", "0,01","N/A", "#N/A"]
platne_meny = ["USD", "EUR", "CZK", "GBP"]
platne_jednotky = ["EA", "FT", "HEC", "IN", "KG", "M", "M2", "MIL", "ROL", "SHT"]

# Rates
eur_usd = 1.16
eur_czk = 24.36
gbp_czk = 27.09
usd_czk = 21
eur_gbp = 0.899

# Untis
hec = 100
mil = 1000

for item in databaze_itemu:
    # purchase data polozky:
    dodavatel = databaze_itemu.get(item).get("Dodavatel")
    nakupcni_cena = databaze_itemu.get(item).get("Purchase price")
    nakupcni_mena = databaze_itemu.get(item).get("Purchase currency")
    nakupni_jednotky = databaze_itemu.get(item).get("Purchase price unit")
    nizka_marze = False


    # sales data polozky:
    sales_cena = databaze_itemu.get(item).get("Actual Sales price now [EUR]")
    
    margin = "TBD"
    databaze_itemu[item]["Actual margin type"] = "N/A"
    
    # Kontrola + spocitani marze pro ciste nakupovane polozky.
    if dodavatel.upper() not in neplatni_dodavatele:
        if nakupcni_cena not in neplatna_cena:
            if nakupcni_mena.upper() in platne_meny:
                if sales_cena.upper() not in neplatna_cena:
                    if nakupni_jednotky.upper() in platne_jednotky:                   
                        
                        if nakupni_jednotky.upper() == "HEC":
                            qty_factor = hec
                        elif nakupni_jednotky.upper() == "MIL":    
                            qty_factor = mil
                        else:
                            qty_factor = 1   

                        if nakupcni_mena.upper() == "EUR":
                           currency_rate = 1
                        elif nakupcni_mena.upper() == "USD":
                            currency_rate = eur_usd
                        elif nakupcni_mena.upper() == "CZK":
                            currency_rate = eur_czk                           
                        elif nakupcni_mena.upper() == "GBP":
                            currency_rate = eur_gbp   

                        # Samotny vypocet marze.
                        margin = (float(sales_cena.replace(",",".")) - (float(nakupcni_cena.replace(",",".")) / qty_factor / currency_rate)) / float(sales_cena.replace(",","."))
                        # Informace o typu marze Purchased.    
                        databaze_itemu[item]["Actual margin type"] = "Purchased"
                        # Informace zda je marze prislis mala.
                        if margin < nizka_marze_hranice:
                            nizka_marze = True
                    else:
                        margin = "Neplatne nakupni jednotky nakupovane polozky."
                else:
                    margin = "Neplatna sales price nakupovane polozky."            
            else:
                margin = "Neplatna purchase mena nakupovane polozky."
        else:
            margin = "Neplatna purchase price nakupovane polozky."    
    
# 10.4.2 Manufactured marze.
    # Pokud neni platny dodavatel, pokusi se spocitat marzi jako pro vyrabenou polozku.
    else:
        if sales_cena.upper() not in neplatna_cena:        
            # Costing data polozky pod nacenovacim projektem.
            costing_nacenovanci_projekt = databaze_itemu.get(item).get("costing NACENOVACI projekt EUR")
            if costing_nacenovanci_projekt.upper() in neplatna_cena:
                costing_nacenovanci_projekt = str(0)
            
            # Costing data polozky prumer projektu last 3 years.
            costing_pmp_last_3_years_avg = databaze_itemu.get(item).get("costing PMP last 3 roky")
            if costing_pmp_last_3_years_avg.upper() in neplatna_cena:
                costing_pmp_last_3_years_avg = str(0)

            # Urceni platneho costingu pro vypocet marze (nacenovaci projekt ma prednost).
            if float(costing_nacenovanci_projekt.replace(",",".")) > 0:
                manufactured_costing_for_margin = costing_nacenovanci_projekt
            elif float(costing_pmp_last_3_years_avg.replace(",",".")) > 0:
                manufactured_costing_for_margin = costing_pmp_last_3_years_avg
            else:
                manufactured_costing_for_margin = "N/A"

            # Pokud existuje platna sales price a costing podle podminek vyse, spocita se Manufactured marze.
            if manufactured_costing_for_margin != "N/A":          
                # Samotny vypocet marze.
                margin = (float(sales_cena.replace(",",".")) - (float(manufactured_costing_for_margin.replace(",",".")))) / float(sales_cena.replace(",","."))
                # Informace o typu marze Marze. 
                databaze_itemu[item]["Actual margin type"] = "Manufactured"
                # Informace zda je marze prislis mala.
                if margin < nizka_marze_hranice:
                    nizka_marze = True
            else:
                margin = "Neplatne costy manufactured polozky."
        else:
            margin = "Neplatna sales price manufactured polozky."                    
    
    # Doplneni informace o marzi do dat. 
    if type(margin) == float:
        databaze_itemu[item]["Actual margin"] = str(margin).replace(".",",")
    else:
        databaze_itemu[item]["Actual margin"] = margin
    # Doplneni informace o nizke marzi do dat.
    if nizka_marze:
        databaze_itemu[item]["Low margin"] = "Pozor - nizka marze"
    else:
        databaze_itemu[item]["Low margin"] = "-"

# 11 Vybrani polozek do vseobecneho BFE katalogu 2023.

# 11.05 Stanoveni vyse eskalaci.
# 11.05.1 # Vseobecna BFE eskalace.

# 9.95 % (Puvodni verze)
# bfe_vyse_eskalace = 0.095

# 11.75 % (Platna verze od 20.9.2022 - mail od Nicolase SCS 2023 CATALOGUES : PROGRESS UPDATE AND FINAL ESCALATION RATES)
bfe_vyse_eskalace = 0.1175

# Jakou hodnotou nasobit lonskou cenu.
eskalacni_faktor_bfe_vseobecne = 1 + bfe_vyse_eskalace

# 11.1 Na zaklade Sales historty z poslednich 3 let.
zahlavi.append("To include (sales history)")

for item in databaze_itemu:
    # Ziskani informace o prodeji daneho itemu "Yes" / "No".
    sales_2020 = databaze_itemu.get(item).get("2020 Sales history orededed Qty")
    sales_2021 = databaze_itemu.get(item).get("2021 Sales history orededed Qty")
    sales_2022 = databaze_itemu.get(item).get("2022 Sales history orededed Qty")

    sales_last_3_years = [sales_2020, sales_2021, sales_2022]

    for sales_in_year in sales_last_3_years:    
        if "YES" in sales_in_year.upper():
            databaze_itemu[item]["To include (sales history)"] = "yes"
            break
    else:
        databaze_itemu[item]["To include (sales history)"] = "no"

# 11.2 Vyber podle Sales history, actual Sales price, not low margin, comments and SG group.
zahlavi.append("To include in 2023 BFE catalogue (sales history, low margin, comment, actual sales price not POR/N/A, correct SG)")
for item in databaze_itemu:
    duvody_por = list()

    comments_ok = False
    actual_price_ok = False
    low_margin_ok = False
    sales_history_ok = False
    correct_sg = False
    
    if "YES" in databaze_itemu.get(item).get("Do katalogu na zaklade commentu z pr. nac. a historicalu").upper():
        comments_ok = True
    else:
        duvody_por.append(f'POR na zaklade commentu z historicalu/pr. nac')
    
    if "N/A" not in databaze_itemu.get(item).get("Actual Sales price now [EUR]").upper():
        actual_price_ok = True
    else:
        duvody_por.append(f'POR na zaklade neplatne Sales price 2022 EUR')
    
    if "POZOR - NIZKA MARZE" not in databaze_itemu.get(item).get("Low margin").upper():
        low_margin_ok = True
    else:
        duvody_por.append(f'POR na zaklade low margin')
    
    if "YES" in databaze_itemu.get(item).get("To include (sales history)").upper():
        sales_history_ok = True
    else:
        duvody_por.append(f'POR na zaklade neplatne sales history')    
    
    # Sales group pravidla.
    sales_group_now = databaze_itemu.get(item).get("Sales group LN").upper()
    sales_group_should_be = databaze_itemu.get(item).get("Sales group to be").upper()
    # SFE dat na POR.
    sfe_sg_ok = False
    if "SFE" not in sales_group_should_be:
        sfe_sg_ok = True
    else:
        duvody_por.append(f'POR na zaklade SFE sales group')
    
    ## Pokud ted je BFE a ma byt MIX
    bfe_to_be_mix = False
    man_placard_bfe_mix_por_nok = False

    if "BFE" in sales_group_now and "MIX" in sales_group_should_be:
        # PLACARDY → POR + MIX do LN.        
        
        if je_to_man_placard(item, databaze_itemu) or je_to_id_placard(item, databaze_itemu):
            man_placard_bfe_mix_por_nok = True
            duvody_por.append(f'POR na zaklade MAN PLACARD BFE → MIX')
        # Jinak ostatni polozky dat jako POR a SG to be MIX.
        else:
            bfe_to_be_mix = True
            duvody_por.append(f'POR na zaklade Sales group BFE ,ale mela by byt MIX')

    # print(item, comments_ok, actual_price_ok, low_margin_ok, sales_history_ok, correct_sg)

    if comments_ok and actual_price_ok and low_margin_ok and sales_history_ok and sfe_sg_ok and not man_placard_bfe_mix_por_nok and not bfe_to_be_mix:
        databaze_itemu[item]["To include in 2023 BFE catalogue (sales history, low margin, comment, actual sales price not POR/N/A, correct SG)"] = "yes"
    else:
        databaze_itemu[item]["To include in 2023 BFE catalogue (sales history, low margin, comment, actual sales price not POR/N/A, correct SG)"] = f'no - {", ".join(duvody_por)}'

# 11.3 Stanoveni BFE ceny 2023 na zaklade vyse BFE eskalace.
zahlavi.append("Vseobecna BFE price 2023 [EUR]")
zahlavi.append("Vseobecna BFE price 2023 [USD]")

man_bfe_placard_2023_price_eur = 57
man_bfe_id_placard_2023_price_eur = 126.03


for item in databaze_itemu:
    price_2023_eur = "?"    
    # print(item)
    # Cenu urcovat jen pro polozky, ktere pujdou do 2023 katalogu.   
    if databaze_itemu.get(item).get("To include in 2023 BFE catalogue (sales history, low margin, comment, actual sales price not POR/N/A, correct SG)").upper() == "YES":
        # Ziskani kurzu USD/EUR na zaklade, kdy byla polozka historicky nacenena.
        usd_eur_ratio = float(databaze_itemu.get(item).get("prepocet USD/EUR podle roku naceneni").replace(",","."))
               
        # Pokud se jedna o MAN PLACARDY nebo ID PLACARDY → jednotnou cenu za placard.
        man_placard = je_to_man_placard(item, databaze_itemu)
        id_placard = je_to_id_placard(item, databaze_itemu)
        
        if man_placard:
            if id_placard:
                databaze_itemu[item]["Vseobecna BFE price 2023 [EUR]"] = str(man_bfe_id_placard_2023_price_eur).replace(".",",")
                databaze_itemu[item]["Vseobecna BFE price 2023 [USD]"] = str(man_bfe_id_placard_2023_price_eur * eur_usd).replace(".",",")
            else:
                databaze_itemu[item]["Vseobecna BFE price 2023 [EUR]"] = str(man_bfe_placard_2023_price_eur).replace(".",",")
                databaze_itemu[item]["Vseobecna BFE price 2023 [USD]"] = str(man_bfe_placard_2023_price_eur * eur_usd).replace(".",",")   
        else:    

            # Pokud se jedna o polozky, ktere jsou i v SFE AB katalogu pro 2023 (MIX), dat stejnou USD a EUR cenu jako maji v SFE katalogu (chceme aby cena byla stejna).
            if databaze_itemu.get(item).get("SFE 2023 katalog poslano do Francie 100 % cena [USD]") != "N/A": # Kontrola, zda je item i v SFE katalogu (nema N/A).
                # Rovnou se da USD cena jako kopie udaje z SFE katalogu.
                sfe_standard_price_usd = float(databaze_itemu.get(item).get("SFE 2023 katalog poslano do Francie 100 % cena [USD]").replace(",","."))
                # Eur cena se dopocita pomoci usd/eur kurzu.
                eur_price_calculated = sfe_standard_price_usd / usd_eur_ratio

                databaze_itemu[item]["Vseobecna BFE price 2023 [EUR]"] = str(eur_price_calculated).replace(".",",")
                databaze_itemu[item]["Vseobecna BFE price 2023 [USD]"] = str(sfe_standard_price_usd).replace(".",",")
                
            
            # Jinak se cena spocita klasicky pro BFE polozku.
            else:
                # print(item)
                # Ziskani EUR ceny 2022
                price_2022_eur = float(databaze_itemu.get(item).get("Actual Sales price now [EUR]").replace(",","."))

                # Ziskani USD ceny 2022
                # Pokud neni cena v historicalu, mela by byt v prubeznem naceneni.
                if "N/A" in databaze_itemu.get(item).get("BFE historical 2022 [USD]"):
                    if "N/A" in databaze_itemu.get(item).get("BFE pr. Nac 2022"):
                        price_2022_usd = price_2022_eur * usd_eur_ratio 
                    else:
                        price_2022_usd = float(databaze_itemu.get(item).get("BFE pr. Nac 2022").replace(",",".")) * usd_eur_ratio 
                else:            
                    price_2022_usd = float(databaze_itemu.get(item).get("BFE historical 2022 [USD]").replace(",","."))

                # 2023 EUR cena pocitana jako EUR 2022 cena * (1 + eskalacni procento).
                price_2023_eur = round(price_2022_eur * eskalacni_faktor_bfe_vseobecne, 2)
                # 2023 USD cena pocitana jako 2022 USD cena * (1 + eskalacni procento).
                price_2023_usd = round(price_2022_usd * eskalacni_faktor_bfe_vseobecne, 2)
                
                databaze_itemu[item]["Vseobecna BFE price 2023 [EUR]"] = str(price_2023_eur).replace(".",",")
                databaze_itemu[item]["Vseobecna BFE price 2023 [USD]"] = str(price_2023_usd).replace(".",",")
    else:
        databaze_itemu[item]["Vseobecna BFE price 2023 [EUR]"] = "POR"
        databaze_itemu[item]["Vseobecna BFE price 2023 [USD]"] = "POR"

## 12 Vybrani polozek do zakaznickych ceniku 2023 + stanoveni cen v zakaznickych cenicich.

# 12.1 KLM
# 12.1.1  Vybrat do ceniku vsechny jejich SCOPE A NON SCOPE polozky + BFE polozky z vseobecneho ceniku prodane jim posledni rok. Nedavat BFE polozky bez prodeju podlesni 3 roky.
zahlavi.append("KLM items to include.")
for item in databaze_itemu:
    klm_class = databaze_itemu.get(item).get("KLM SCOPE/NONSCOPE/BFE").upper()
    klm_sales_history = databaze_itemu.get(item).get("KLM sales history 1 rok").upper()
    
    # print(item, klm_class, klm_sales_history)
    # 1. Automaticky pridat SCOPE a NONSCOPE polozky.
    if klm_class == "S" or klm_class == "NS":
        if klm_class == "S":
            databaze_itemu[item]["KLM items to include."] = "yes - SCOPE"
        else:
            databaze_itemu[item]["KLM items to include."] = "yes - NONSCOPE"           
    # 2. Jinak pokud je to polozka, kterou meli uz letos v ceniku, pridat, pokud bude i v BFE vseobecnem ceniku.
    elif klm_class == "BFE" and databaze_itemu.get(item).get("To include in 2023 BFE catalogue (sales history, low margin, comment, actual sales price not POR/N/A, correct SG)").upper() == "YES":
        databaze_itemu[item]["KLM items to include."] = "yes - BFE"
    # 3. Pokud je to polozka, ktera se jim prodala tento rok a bude i v BFE vseobecnem → pridat.
    elif klm_sales_history == "YES" and databaze_itemu.get(item).get("To include in 2023 BFE catalogue (sales history, low margin, comment, actual sales price not POR/N/A, correct SG)").upper() == "YES":
        databaze_itemu[item]["KLM items to include."] = "yes - BFE"
    else:
        databaze_itemu[item]["KLM items to include."] = "no" 

# 12.1.2  Stanovit cenu pro dalsi rok. BFE vseobecnou cenu pro BFE polozky a Scope/Nonscope osalovane ceny pro dalsi rok.
zahlavi.append("KLM items 2023 price [EUR]")

# KLM Scope + NONscope esklace pro 2023 = +3 % (9.504 % → CAP 3 %)
klm_eskalace = 0.03
klm_eskalaceni_faktor = 1 + klm_eskalace

for item in databaze_itemu:
    klm_item_status = databaze_itemu.get(item).get("KLM items to include.")
    
    # N/A cena pro polozky, ktere nebudou v jejich ceniku.
    if klm_item_status == "no":
        databaze_itemu[item]["KLM items 2023 price [EUR]"] = "N/A" 
    # Pokud se jedna o BFE polozku, ktera pujde jim do ceniku a neni SCOPE / NONSCOPE → dat BFE vseobecnou cenu EUR.
    elif klm_item_status == "yes - BFE":
        databaze_itemu[item]["KLM items 2023 price [EUR]"] = databaze_itemu.get(item).get("Vseobecna BFE price 2023 [EUR]")
    # Pokud se jedna o jejich SCOPE/NONSCOPE polozku → stanovit novou cenu jako EUR cenu predchoziho roku * KLM eskalace.
    elif klm_item_status == "yes - NONSCOPE" or klm_item_status == "yes - SCOPE":
        databaze_itemu[item]["KLM items 2023 price [EUR]"] = str(round(float(databaze_itemu.get(item).get("KLM items 2022 [Eur]").replace(",",".")) * klm_eskalaceni_faktor,2)).replace(".",",")


# Transavia
# 12.2.1  Vybrat vsechny BFE polozky s prodejem k nim za posledni 3 roky, ktere pujdou do BFE vseobecneho + jejich zlevnene polozky co pujdou do BFE vseobecneho.
zahlavi.append("Transavia items to include.")
for item in databaze_itemu:
    transavia_class = databaze_itemu.get(item).get("Transavia zlevnene polozky").upper()
    transavia_sales_history = databaze_itemu.get(item).get("Transavia sales history 3 roky").upper()
    # 1. Pokud se jedna o polozku, ktera pujde do BFE vseobecneho:
    if databaze_itemu.get(item).get("To include in 2023 BFE catalogue (sales history, low margin, comment, actual sales price not POR/N/A, correct SG)").upper() == "YES":
        # 1A. Pokud se jedna o jejich zlevnenou polozku (s) a pujde do vseobecneho ceniku → pridat.
        if transavia_class == "S":
            databaze_itemu[item]["Transavia items to include."] = "yes - zlevnena"     
        # 1B. Jinak, pokud se jim prodala polozka posledni 3 roky a pujde do vseobecneho ceniku → pridat.
        elif transavia_sales_history == "YES":
            databaze_itemu[item]["Transavia items to include."] = "yes - BFE"
        # 1C. Pokud se jim neprodala posledni 3 roky → nepujde do jejich ceniku.
        else:
            databaze_itemu[item]["Transavia items to include."] = "no"                 
    # 2. Polozky ktere nebudou letos v BFE vseobecnem jim nepridavat.
    else:
        if transavia_class == "S":
            databaze_itemu[item]["Transavia items to include."] = "no - zlevnena. Oznacit v BFE historicalu aby se pri dalsim naceneni jim dala pomerne mensi cena jako meli drive."        
        else:    
            databaze_itemu[item]["Transavia items to include."] = "no"

# 12.2.2 Stanovit cenu pro dalsi rok. BFE vseobecnou cenu pro BFE polozky a jejich zlevnenou cenu navysenou o eskalaci u jejich specialne zlevnenych polozek.
zahlavi.append("Transavia items 2023 price [EUR]")
for item in databaze_itemu:
    transavia_item_status = databaze_itemu.get(item).get("Transavia items to include.")

    # N/A cena pro polozky, ktere nebudou v jejich ceniku.
    if transavia_item_status == "no":
        databaze_itemu[item]["Transavia items 2023 price [EUR]"] = "N/A" 
    # Pokud se jedna o BFE polozku, ktera pujde jim do ceniku a neni jeich zlevnena → dat BFE vseobecnou cenu EUR.
    elif transavia_item_status == "yes - BFE":
        databaze_itemu[item]["Transavia items 2023 price [EUR]"] = databaze_itemu.get(item).get("Vseobecna BFE price 2023 [EUR]")
    # Pokud se jedna o jejich zlevnenou polozku → stanovit novou cenu jako EUR cenu predchoziho roku * BFE eskalace.
    elif transavia_item_status == "yes - zlevnena":
        # Pokud se jedna o jejich histroicky zlevnenou polozku, ale momentalne neni pro ne cena stanovena → dat poznamku o zlevnene polozce az se bude znovu nacenovat.
        if "N/A" in databaze_itemu.get(item).get("Transavia items 2022 [Eur]"):
            databaze_itemu[item]["Transavia items 2023 price [EUR]"] = "zlevnena polozka - Az se bude znovu nacenovat, dat pomerne mensi cenu jako meli drive."
        # Pokud mame znamou jejich zlevnenou cenu → tu navysit o BFE eskalaci.
        else:
            databaze_itemu[item]["Transavia items 2023 price [EUR]"] = str(round(float(databaze_itemu.get(item).get("Transavia items 2022 [Eur]").replace(",",".")) * eskalacni_faktor_bfe_vseobecne,2)).replace(".",",")
    # Pokud je to jejich zlevnena ale neni momentalne znama jejich zlevnena cena → dat POR s poznamkou a naceneni jako zlevnena polozka az se bude znovu nacenovat.
    elif transavia_item_status == "no - zlevnena. Oznacit v BFE historicalu aby se pri dalsim naceneni jim dala pomerne mensi cena jako meli drive.":
        databaze_itemu[item]["Transavia items 2023 price [EUR]"] = "zlevnena polozka - Az se bude znovu nacenovat, dat pomerne mensi cenu jako meli drive."

# QATAR
# 12.3.1 Vybrat vsechny BFE polozky ktere pujdou do BFE vseobecneho, pote vyradit vsechny polozky, s navysenim ceny vyssim nez STD eskalace + pridat jejich zlevnene polozky, pokud se jim prodaly posledni 3 roky.
zahlavi.append("QATAR items to include.")
for item in databaze_itemu:
    qatar_class = databaze_itemu.get(item).get("QATAR zlevnene polozky").upper()
    qatar_sales_history = databaze_itemu.get(item).get("QATAR sales history 3 roky").upper()
    # 1. Pokud se jedna o jejich zlevnenou polozku (s). a prodavala se jim podledni 3 roky → pridat, pokud nepridavala → oznacit v historicalu ze je jejich zlevnena.
    if qatar_class == "S":
        # Pokud pujde i do vseobecneho BFE ceniku 2023 → pridat jim ji jako zlevnenou.
        if databaze_itemu.get(item).get("To include in 2023 BFE catalogue (sales history, low margin, comment, actual sales price not POR/N/A, correct SG)").upper() == "YES":
            databaze_itemu[item]["QATAR items to include."] = "yes - zlevnena" 
        # Jinak pokud se jim prodavala posledni 3 roky → pridat jak zlevnenou.
        elif qatar_sales_history == "YES":                           
            databaze_itemu[item]["QATAR items to include."] = "yes - zlevnena"        
        # Pokud se jejich zlevnena jim neprodala posledni 3 roky, ani nepujde do vseobecneho ceniku BFE 2023 → dat ji jako POR s informaci ze se jedna o zlevnenou polozku.
        else:
            databaze_itemu[item]["QATAR items to include."] = "no - zlevnena. Oznacit v BFE historicalu aby se pri dalsim naceneni jim dala pomerne mensi cena jako meli drive."    
    # 2. Pokud to neni jejich zlevnena polozka a jedna se o polozku, ktera pujde do BFE vseobecneho → pridat, jinak nepridavat:
    else:    
        if databaze_itemu.get(item).get("To include in 2023 BFE catalogue (sales history, low margin, comment, actual sales price not POR/N/A, correct SG)").upper() == "YES":
            databaze_itemu[item]["QATAR items to include."] = "yes - BFE"
        else:
            databaze_itemu[item]["QATAR items to include."] = "no"

# 12.3.1 Stanovit cenu pro dalsi rok. BFE vseobecnou cenu pro BFE polozky a jejich zlevnenou cenu navysenou o eskalaci u jejich specialne zlevnenych polozek.
zahlavi.append("QATAR items 2023 price [EUR]")
for item in databaze_itemu:
    qatar_item_status = databaze_itemu.get(item).get("QATAR items to include.")

    # N/A cena pro polozky, ktere nebudou v jejich ceniku.
    if qatar_item_status == "no":
        databaze_itemu[item]["QATAR items 2023 price [EUR]"] = "N/A" 
    # Pokud se jedna o BFE polozku, ktera pujde jim do ceniku a neni jeich zlevnena → dat BFE vseobecnou cenu EUR.
    elif qatar_item_status == "yes - BFE":
        databaze_itemu[item]["QATAR items 2023 price [EUR]"] = databaze_itemu.get(item).get("Vseobecna BFE price 2023 [EUR]")
    # Pokud se jedna o jejich zlevnenou polozku → stanovit novou cenu jako EUR cenu predchoziho roku * BFE eskalace.
    elif qatar_item_status == "yes - zlevnena":
        # Pokud se jedna o jejich histroicky zlevnenou polozku, ale momentalne neni pro ne cena stanovena → dat poznamku o zlevnene polozce az se bude znovu nacenovat.
        if "N/A" in databaze_itemu.get(item).get("QATAR items 2022 [Eur]"):
            databaze_itemu[item]["QATAR items 2023 price [EUR]"] = "zlevnena polozka - Az se bude znovu nacenovat, dat pomerne mensi cenu jako meli drive."
        # Pokud mame znamou jejich zlevnenou cenu → tu navysit o BFE eskalaci.
        else:
            databaze_itemu[item]["QATAR items 2023 price [EUR]"] = str(round(float(databaze_itemu.get(item).get("QATAR items 2022 [Eur]").replace(",",".")) * eskalacni_faktor_bfe_vseobecne,2)).replace(".",",")
    # Pokud je to jejich zlevnena ale neni momentalne znama jejich zlevnena cena → dat POR s poznamkou a naceneni jako zlevnena polozka az se bude znovu nacenovat.
    elif qatar_item_status == "no - zlevnena. Oznacit v BFE historicalu aby se pri dalsim naceneni jim dala pomerne mensi cena jako meli drive.":
        databaze_itemu[item]["QATAR items 2023 price [EUR]"] = "zlevnena polozka - Az se bude znovu nacenovat, dat pomerne mensi cenu jako meli drive."


output_to_write = list()
output_to_write.append(zahlavi)

for item in databaze_itemu:    
    line_to_to_write = list()    
    for key in databaze_itemu.get(item):
        line_to_to_write.append(databaze_itemu[item][key])
    output_to_write.append(line_to_to_write)

with open("BFE output.txt", "w", encoding='utf-8') as output:
    for line in output_to_write:    
        to_write = "\t".join(line)
        output.write(f'{to_write}\n')
    output.close()