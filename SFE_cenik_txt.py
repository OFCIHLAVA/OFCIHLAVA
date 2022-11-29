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

input_file = "Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\6_Tvorba ceníků\\SFE ceníky\\ceniky\\SFE file_txt.txt"

with open(input_file, "r", encoding='utf-8') as file:
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
print(f'Zahlavi: {zahlavi}\n{len(zahlavi)}')


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

print(f'DATABAZE ITEMU VYTVORENA OK 1')

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

print(f'INFO MAN PLACARD / ID PLACARD OK 2')

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
    elif "LAV" in databaze_itemu.get(item).get("Sales group LN").upper():
        sales_group_to_be = "LAV"
    
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

print(f'SALES GROUP OK 3')

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
                    if databaze_itemu.get(item).get("Dodavatel").upper() == "0":                    
                        if "BFE" in databaze_itemu.get(item).get("Sales group to be").upper():
                            databaze_itemu[item]["Sales group to be"] = "BFE02"
                        else:
                            databaze_itemu[item]["Sales group to be"] = "SFE02"
                    else:
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

print(f'01 / 02 SG SKUPINA OK 4')

# 10.2 Sales LT.
zahlavi.append("Sales LT to be")

for item in databaze_itemu:
    sales_lt_ln = databaze_itemu.get(item).get("LN Sales LT [days]")
    sales_lt_python = databaze_itemu.get(item).get("LN Sales LT [days] Python")
    sales_lt_to_be = 0
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
    # Pokud neni udaj ze sales dat → dat co rika Pytohn.
    elif sales_lt_ln == "N/A":
        if sales_lt_python != "N/A":            
            sales_lt_to_be = sales_lt_python
        else:
            sales_lt_to_be = "N/A" 
    # Pokud je sales LN v LN pod 14 dni → dat co je v pythonu.
    elif float(sales_lt_ln) <14:
        sales_lt_to_be = round(float(sales_lt_python.replace(",",".")),0)
    # U vsech ostatnich porovnat lt v LN s Pythonem a dat z Pythonu pokud je vyssi nez z LN, jinak nechat LN LT.
    else:
        if float(sales_lt_ln.replace(",",".")) < float(sales_lt_python.replace(",",".")):
            sales_lt_to_be = round(float(sales_lt_python.replace(",",".")),0)
        else:
            sales_lt_to_be = sales_lt_ln

    # Pro N/A hodnoty dat do LN 0 (Da sales LT lze zadat pouze ciselnou hodnotu)
    if sales_lt_to_be == "N/A":
        sales_lt_to_be = 0    
    
    databaze_itemu[item]["Sales LT to be"] = str(sales_lt_to_be).replace(".",",")

print(f'SALES LT TO BE OK 5')


# 10.3 Sales price actual now.
zahlavi.append("Actual Sales price now [USD]")

for item in databaze_itemu:
    historical_usd_price = databaze_itemu.get(item).get("SFE historical 2022 100%")
    pr_nac_usd_price = databaze_itemu.get(item).get("SFE pr. Nac 2022")

    # Pokud je platna cena z prubezneho nacenovani → cena z pr nac.
    if pr_nac_usd_price != "0" and "N/A" not in pr_nac_usd_price.upper():
        databaze_itemu[item]["Actual Sales price now [USD]"] = pr_nac_usd_price.replace(" ","")
    else:
        # Pokud je platna cena z historicalu → cena z historicalu.
        if historical_usd_price != "0" and "N/A" not in historical_usd_price.upper() and "PRICE ON REQUEST" not in historical_usd_price.upper() and "BFE" not in historical_usd_price.upper():
            databaze_itemu[item]["Actual Sales price now [USD]"] = historical_usd_price.replace(" ","")
        # Pokud ani jedno z vysse neplati → "N/A"
        else:
            databaze_itemu[item]["Actual Sales price now [USD]"] = "N/A"

print(f'ACTUAL SALES PRICE OK 6')

# 10.3.5 MOQ to be.

# 2023 vyse eskalace = 4,71 %
sfe_vyse_eskalace = 0.0471
eskalacni_faktor_sfe = 1 + sfe_vyse_eskalace

zahlavi.append("Sales MOQ to be")

for item in databaze_itemu:
    sfe_sales_price_now = "N/A"
    sfe_sales_price_to_be = "N/A"
    moq_now = 'N/A'
    moq_to_be = 'N/A'

    # sales price s eskalacei, aby se dalo spocitat MOQ na dalsi rok podle cenovzch pravidel
    if databaze_itemu.get(item).get("Actual Sales price now [USD]") != "N/A":
        sfe_sales_price_now = float(databaze_itemu.get(item).get("Actual Sales price now [USD]").replace(",","."))
        sfe_sales_price_to_be = round(sfe_sales_price_now * eskalacni_faktor_sfe,2)

    # MOQ ted, abz se dalo zkontrolovat spoecialni MOQ na dilech typu sheety / decory atp., kde neni celocislene MOQ. 
    if "N/A" not in databaze_itemu.get(item).get("MOQ LN"):
        moq_now = float(databaze_itemu.get(item).get("MOQ LN").replace(",",".")) 
  
    # 1. MAN PLACARDY = 3, ID placardy = 1
    man_placard = je_to_man_placard(item, databaze_itemu)
    id_placard = je_to_id_placard(item, databaze_itemu)
    
    if man_placard:
        if id_placard:
            moq_to_be = 1
        else:
            moq_to_be = 4
    
    # 2. Pokud neni MOQ neplatne (N/A), proverit, zda se jedna o zvlastni necele MOQ a pripadbne proverit, jake ma byt MOQ.
    else:        
        # 2.1 Pro "divne MOQ vyhodit info k dalsimu provereni jake je spravne MOQ (decory, sheety atp.)"
        if moq_now != 'N/A':
            if moq_now % 1 != 0:
                moq_to_be = "POZOR, divne MOQ - proverit s Petrem, jake dat MOQ."
        
            # 2.2 Pro "nedivne MOQ: Pokud je znama cena na dalsi rok → dat MOQ podle tabulky z SFE pravidel MOQ, jinak pokud neni znam → dat "N/A".
            else:
                # Pokud je znama cena na pristi rok:
                if sfe_sales_price_to_be != "N/A":
                    # a) Cena to be vyssi nez 100 USD → MOQ 1
                    if sfe_sales_price_to_be > 100:
                        moq_to_be = 1
                    elif sfe_sales_price_to_be >=50:    
                        moq_to_be = 2
                    elif sfe_sales_price_to_be >=30:    
                        moq_to_be = 4
                    elif sfe_sales_price_to_be >=11:    
                        moq_to_be = 10
                    elif sfe_sales_price_to_be >0:    
                        moq_to_be = 25
                    elif sfe_sales_price_to_be >=30:    
                        moq_to_be = f'Pozor, cen na dalsi rok <= 0 !'
                else:
                    moq_to_be = f'Pozor, neznama cena na dalsi rok → neni mozno urcit MOQ → dat 0 do LN dat na pristi rok.'
            # print(f'{item}, {sfe_sales_price_to_be}, {moq_to_be}')    
        else:
            f'Pozor, nezname MOQ v LN sales datech → neni mozno urcit MOQ → dat 0 do LN dat na pristi rok.'          
    databaze_itemu[item]["Sales MOQ to be"] = str(moq_to_be)
    
print(f'MOQ OK 8')


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
    sales_cena = databaze_itemu.get(item).get("Actual Sales price now [USD]")
    
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
                        # print(f'{item}, {sales_cena}, {nakupcni_cena}')
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

print(f'MARZE OK 9')

 
# 11 Stanoveni vyse eskalaci.
# 11.05.1 # SFE eskalace.

# 4.71 % pro 2023

# Uz pocitao v MOQ sekci:
## # 2023 vyse eskalace = 4,71 %
## sfe_vyse_eskalace = 0.0471
## eskalacni_faktor_sfe = 1 + sfe_vyse_eskalace

# Jakou hodnotou nasobit lonskou cenu.
eskalacni_faktor_sfe = eskalacni_faktor_sfe # Hodnota jen pravzata z sekce MOQ, kde uz byla stanovena.  

# 12 Vybrani polozek do SFE AIRBUS katalogu 2023 a stanoveni ceny pro 2023.

    # A. Polozky, ktere jsou letos v AB katalogu a maji SServices Sales history last 3 years.
    # B. Polozky ze seznamu RFQ ze SServices, ktere maji Sales history posledni rok (SServices SH + CZ SH).
    # C. Samotne vybrani polozek podle podminek vyse + doplnujicich kritetii (low margin, komentare z historicalu / pr nac., actual sales price not POR/N/A, correct SG))
    # D. Stanoveni ceny na dalsi rok podle vyse sfe eskalace

# 12.A Na zaklade Letos v katalogu a Sales history SServices z poslednich 3 let.

zahlavi.append("To include AB katalog (sales history)")
 
for item in databaze_itemu:
    # Ziskani informace o service prodeji daneho itemu "Yes" / "No".
    services_sh_2020 = databaze_itemu.get(item).get("2020 SS Sales history orededed Qty")
    services_sh_2021 = databaze_itemu.get(item).get("2021 SS Sales history orededed Qty")
    services_sh_2022 = databaze_itemu.get(item).get("2022 SS Sales history orededed Qty")
 
    services_sh_last_3_years = [services_sh_2020, services_sh_2021, services_sh_2022]

    # Ziskani informace, zda je v soucasnosti v AB katalogu.
    letos_v_katalogu = databaze_itemu.get(item).get("AIRBUS 2022 katalog [USD]")

    # Pokud je to polozka letos v AB katalogu:
    if letos_v_katalogu != "N/A":
        # Pokud ma services prodeje posledni 3 roky → include.
        for sh_in_year in services_sh_last_3_years:    
            if "YES" in sh_in_year.upper():
                databaze_itemu[item]["To include AB katalog (sales history)"] = "yes"
                break
        else:
            databaze_itemu[item]["To include AB katalog (sales history)"] = "no"                
    # Jinak nepridavat
    else:
        databaze_itemu[item]["To include AB katalog (sales history)"] = "no"

print(f'AB INCLUDE SH OK 10')


# 12.B Pridat polozky z RFQ seznamu od SServices, ktere maji letosni a lonskou SH services / SH CZ.
zahlavi.append("To include AB katalog (RFQ with CZ sales history)")
for item in databaze_itemu:
    
    # Ziskani informace zda se jedna o polozku poptavanou skrze RFQ posledni rok.
    rfq_polozka = "N/A"
    rfq_polozka = databaze_itemu.get(item).get("RFQ 2021-2022 SS Sales history")

    # Ziskani informace o CZ prodeji daneho itemu "Yes" / "No".
    cz_sh_2020 = databaze_itemu.get(item).get("2020 CZ Sales history orededed Qty")
    cz_sh_2021 = databaze_itemu.get(item).get("2021 CZ Sales history orededed Qty")
    cz_sh_2022 = databaze_itemu.get(item).get("2022 CZ Sales history orededed Qty")
 
    cz_sh_this_year_and_last_years = [cz_sh_2021, cz_sh_2022]

    # Pokud je polozka z RFQ seznamu:
    if rfq_polozka.upper() == "YES":
        # Pokud ma CZ sales history letos / loni → include.
        for sh_in_year in cz_sh_this_year_and_last_years:    
            if "YES" in sh_in_year.upper():
                databaze_itemu[item]["To include AB katalog (RFQ with CZ sales history)"] = "yes"
                break
            else:
                databaze_itemu[item]["To include AB katalog (RFQ with CZ sales history)"] = "no"    
    else:
        databaze_itemu[item]["To include AB katalog (RFQ with CZ sales history)"] = "no" 

print(f'AB RFQ WITH SH OK 11')

# 12.C Samotny vyber polozek da AB katalogu 2023:
zahlavi.append("To include in 2023 AIRBUS catalogue (sales history, low margin, comment, actual sales price not POR/N/A, RFQ sales history, correct SG)")
for item in databaze_itemu:   
    duvody_por = list()

    # Jednotlive testy:
    comments_ok = False
    actual_price_ok = False
    low_margin_ok = False
    sales_history_ok = False

    sfe_sg_ok = False
    blk_sg_ok = False
    bfe_to_be_mix = False
    man_placard_bfe_mix_por_nok = False
    correct_sg = False

    # Comments check:    
    if "YES" in databaze_itemu.get(item).get("Do AIRBUS katalogu na zaklade commentu z pr. nac. a historicalu").upper():
        comments_ok = True
    else:
        duvody_por.append(f'POR na zaklade commentu z historicalu/pr. nac')
    # Platna cena check:      
    if "N/A" not in databaze_itemu.get(item).get("Actual Sales price now [USD]").upper():
        actual_price_ok = True
    else:
        duvody_por.append(f'POR na zaklade neplatne Sales price 2022 USD')
    # Low margin check:    
    if "POZOR - NIZKA MARZE" not in databaze_itemu.get(item).get("Low margin").upper():
        low_margin_ok = True
    else:
        duvody_por.append(f'POR na zaklade low margin')
    # Kat. sales history + RFQ sales history check:      
    if "YES" in databaze_itemu.get(item).get("To include AB katalog (sales history)").upper():
        sales_history_ok = True
    elif "YES" in databaze_itemu.get(item).get("To include AB katalog (RFQ with CZ sales history)").upper():
        sales_history_ok = True
    else:
        duvody_por.append(f'POR na zaklade neplatne sales history')    

    ### Sales group pravidla check:
    sales_group_now = databaze_itemu.get(item).get("Sales group LN").upper()
    sales_group_should_be = databaze_itemu.get(item).get("Sales group to be").upper()

    # BFE dat na POR.
    if "BFE" not in sales_group_should_be:
        sfe_sg_ok = True
    else:
        duvody_por.append(f'POR na zaklade BFE sales group')

    # BLK dat na POR.
    if "BLK" not in sales_group_should_be:
        blk_sg_ok = True
    else:
        duvody_por.append(f'POR na zaklade BLK sales group')

    # Pokud ted je BFE a ma byt MIX

    if "BFE" in sales_group_now and "MIX" in sales_group_should_be:
        # PLACARDY → POR + MIX do LN.        
        
        if je_to_man_placard(item, databaze_itemu) or je_to_id_placard(item, databaze_itemu):
            man_placard_bfe_mix_por_nok = True
            duvody_por.append(f'POR na zaklade MAN PLACARD BFE → MIX')
        # Jinak ostatni polozky dat jako POR a SG to be MIX.
        else:
            bfe_to_be_mix = True
            duvody_por.append(f'POR na zaklade Sales group BFE ,ale mela by byt MIX')


    # Celkovy SG check:
    if sfe_sg_ok and blk_sg_ok and (not man_placard_bfe_mix_por_nok and not bfe_to_be_mix):
        correct_sg = True
    ###
    # print(item, comments_ok, actual_price_ok, low_margin_ok, sales_history_ok, correct_sg)
    if comments_ok and actual_price_ok and low_margin_ok and sales_history_ok and correct_sg:
        databaze_itemu[item]["To include in 2023 AIRBUS catalogue (sales history, low margin, comment, actual sales price not POR/N/A, RFQ sales history, correct SG)"] = "yes"
    else:
        databaze_itemu[item]["To include in 2023 AIRBUS catalogue (sales history, low margin, comment, actual sales price not POR/N/A, RFQ sales history, correct SG)"] = f'no - {", ".join(duvody_por)}'

print(f'TO INCLUDE IN AIRBUS CATALOGUE OK 12')


# 12.D Stanoveni ceny pro AB katalog polozky na dalsi rok.
# Polozky, ktere pujdou do AV katalogu pronasobit sfe eskalaci na dalsi rok.
zahlavi.append("AIRBUS 2023 STANDARD PRICE TO BE [USD]")

for item in databaze_itemu:
    sfe_sales_price_now = "N/A"
    ab_cat_standard_price_to_be = "N/A"

    # Pokud polozka pujde do AB katalogu → stanovit standard price na dalsi rok.
    if databaze_itemu.get(item).get("To include in 2023 AIRBUS catalogue (sales history, low margin, comment, actual sales price not POR/N/A, RFQ sales history, correct SG)").upper() == "YES":
        sfe_sales_price_now = float(databaze_itemu.get(item).get("Actual Sales price now [USD]").replace(",","."))
        ab_cat_standard_price_to_be = round(sfe_sales_price_now * eskalacni_faktor_sfe,2)
    # Jinak nedavat nic.
    else:
        ab_cat_standard_price_to_be = "-"
    databaze_itemu[item]["AIRBUS 2023 STANDARD PRICE TO BE [USD]"] = str(ab_cat_standard_price_to_be).replace(".",",")

print(f'AB KATALOG STANDARD PRICE OK 13')

# 13. Vybrani polozek do "INTERNAL" ceniku 2023 (maji platnou cenu a neni je treba znovu nacenovat, ale nebudou obsazeny v AB katalogu pristi rok).
# + Ostatni polozky oznaceny v Historicalu jako POR.
# (V podstate to stejne jako v kroku 13, dalo by se sloucit. Rozdeleno pouze pro prehlednost.)

    # A. Polozky, ktere maji CZ Sales history last 3 years.
    # B. Samotne vybrani polozek podle podminek vyse + doplnujicich kritetii (low margin, komentare z historicalu / pr nac., actual sales price not POR/N/A, correct SG))
    # C. Stanoveni ceny na dalsi rok podle vyse sfe eskalace

# 13.A Vybrani polozek do INTERNALU na zaklade CZ Sales history posledni 3 roky:
zahlavi.append("To include INTERNAL katalog (sales history)")
for item in databaze_itemu:
    # Ziskani informace o CZ prodeji daneho itemu "Yes" / "No".
    cz_sh_2020 = databaze_itemu.get(item).get("2020 CZ Sales history orededed Qty")
    cz_sh_2021 = databaze_itemu.get(item).get("2021 CZ Sales history orededed Qty")
    cz_sh_2022 = databaze_itemu.get(item).get("2022 CZ Sales history orededed Qty")
 
    cz_sh_last_3_years = [cz_sh_2020, cz_sh_2021, cz_sh_2022]

    # Pokud ma CZ sales history posledni 3 roky → include.
    for sh_in_year in cz_sh_last_3_years:    
        if "YES" in sh_in_year.upper():
            databaze_itemu[item]["To include INTERNAL katalog (sales history)"] = "yes"
            break
        else:
            databaze_itemu[item]["To include INTERNAL katalog (sales history)"] = "no"

print(f'TO INCLUDE INTERNAL CATALOGUE SH OK 14')

# 13.B Samotne vybrani polozek do INTERNALU:
zahlavi.append("To include in 2023 INTERNAL catalogue (sales history, low margin, comment, actual sales price not POR/N/A, correct SG)")
for item in databaze_itemu:   
    if databaze_itemu.get(item).get("To include in 2023 AIRBUS catalogue (sales history, low margin, comment, actual sales price not POR/N/A, RFQ sales history, correct SG)").upper() != "YES":
        duvody_por = list()

        # Jednotlive testy:
        comments_ok = False
        actual_price_ok = False
        low_margin_ok = False
        sales_history_ok = False

        sfe_sg_ok = False
        blk_sg_ok = False
        bfe_to_be_mix = False
        man_placard_bfe_mix_por_nok = False
        correct_sg = False

        # Comments check:    
        if "YES" in databaze_itemu.get(item).get("Do INTERNAL na zaklade commentu z pr. nac. a historicalu").upper():
            comments_ok = True
        else:
            duvody_por.append(f'POR na zaklade commentu z historicalu/pr. nac')
        # Platna cena check:      
        if "N/A" not in databaze_itemu.get(item).get("Actual Sales price now [USD]").upper():
            actual_price_ok = True
        else:
            duvody_por.append(f'POR na zaklade neplatne Sales price 2022 USD')
        # Low margin check:    
        if "POZOR - NIZKA MARZE" not in databaze_itemu.get(item).get("Low margin").upper():
            low_margin_ok = True
        else:
            duvody_por.append(f'POR na zaklade low margin')
        # Sales history + RFQ sales history check:      
        if "YES" in databaze_itemu.get(item).get("To include INTERNAL katalog (sales history)").upper():
            sales_history_ok = True
        else:
            duvody_por.append(f'POR na zaklade neplatne sales history')    

        ### Sales group pravidla check:
        sales_group_now = databaze_itemu.get(item).get("Sales group LN").upper()
        sales_group_should_be = databaze_itemu.get(item).get("Sales group to be").upper()

        # BFE dat na POR.
        if "BFE" not in sales_group_should_be:
            sfe_sg_ok = True
        else:
            duvody_por.append(f'POR na zaklade BFE sales group')

        # BLK dat na POR.
        if "BLK" not in sales_group_should_be:
            blk_sg_ok = True
        else:
            duvody_por.append(f'POR na zaklade BLK sales group')

        # Pokud ted je BFE a ma byt MIX

        if "BFE" in sales_group_now and "MIX" in sales_group_should_be:
            # PLACARDY → POR + MIX do LN.        
            
            if je_to_man_placard(item, databaze_itemu) or je_to_id_placard(item, databaze_itemu):
                man_placard_bfe_mix_por_nok = True
                duvody_por.append(f'POR na zaklade MAN PLACARD BFE → MIX')
            # Jinak ostatni polozky dat jako POR a SG to be MIX.
            else:
                bfe_to_be_mix = True
                duvody_por.append(f'POR na zaklade Sales group BFE ,ale mela by byt MIX')

        # Celkovy SG check:
        if sfe_sg_ok and blk_sg_ok and (not man_placard_bfe_mix_por_nok and not bfe_to_be_mix):
            correct_sg = True
        ###
        # print(item, comments_ok, actual_price_ok, low_margin_ok, sales_history_ok, correct_sg)
        if comments_ok and actual_price_ok and low_margin_ok and sales_history_ok and correct_sg:
            databaze_itemu[item]["To include in 2023 INTERNAL catalogue (sales history, low margin, comment, actual sales price not POR/N/A, correct SG)"] = "yes"
        else:
            databaze_itemu[item]["To include in 2023 INTERNAL catalogue (sales history, low margin, comment, actual sales price not POR/N/A, correct SG)"] = f'no - {", ".join(duvody_por)}'
    else:
        databaze_itemu[item]["To include in 2023 INTERNAL catalogue (sales history, low margin, comment, actual sales price not POR/N/A, correct SG)"] = "no - v AB katalogu"

print(f'TO INCLUDE IN INTERNAL CATALOGUE OK 15')

# 13.C Stanoveni ceny pro INTERNAL katalog polozky na dalsi rok.
# Polozky, ktere pujdou do INTERNAL katalogu pronasobit sfe eskalaci na dalsi rok.
zahlavi.append("INTERNAL 2023 STANDARD PRICE TO BE [USD]")

for item in databaze_itemu:
    sfe_sales_price_now = "N/A"
    internal_standard_price_to_be = "N/A"
    bfe_kat_price_usd = databaze_itemu.get(item).get("BFE 2023 katalog [USD]")

    # Pokud polozka pujde do internal katalogu → stanovit standard price na dalsi rok.
    if databaze_itemu.get(item).get("To include in 2023 INTERNAL catalogue (sales history, low margin, comment, actual sales price not POR/N/A, correct SG)").upper() == "YES":
        # Pokud polozka ma cenu v BFE USD ceniku, melo by se jednat o MIX polozku a cena by mela byt stejna.
        if bfe_kat_price_usd != "N/A":
            internal_standard_price_to_be = bfe_kat_price_usd
        # Jinak vzit platnou SFE cenu a pronasobit SFE eskalaci.
        else:
            sfe_sales_price_now = float(databaze_itemu.get(item).get("Actual Sales price now [USD]").replace(",","."))
            internal_standard_price_to_be = round(sfe_sales_price_now * eskalacni_faktor_sfe,2)
    # Jinak nedavat nic.
    else:
        internal_standard_price_to_be = "-"
    databaze_itemu[item]["INTERNAL 2023 STANDARD PRICE TO BE [USD]"] = str(internal_standard_price_to_be).replace(".",",")

print(f'INTERNAL KATALOG STANDARD PRICE OK 16')

####################################################################
###### Toto nebude potreba pro dalsi roky, relevantni pouze pro 2023.  ↓↓↓↓↓↓↓↓↓
####################################################################

# 14 Vybrani polozek do "Realneho" Internalu pro 2023.
# Pro rok 2023 se nevytvoril spravne AB katalog → INTERNAL by taku neodpovidal. Nutnu upravit pro 2023 taky. 
zahlavi.append("TO INCLUDE 'Realny' INTERNAL 2023")
for item in databaze_itemu:
    # Ziskani info, zda includovat v Internal.
    include_internal = databaze_itemu.get(item).get("To include in 2023 INTERNAL catalogue (sales history, low margin, comment, actual sales price not POR/N/A, correct SG)")
    # Ziskani info, zda v 'realnem' AB katalogu 2023.
    in_real_ab_cat_2023 = databaze_itemu.get(item).get("AIRBUS 2023 katalog [USD]")

    # Pokud polozka ma 'yes' nebo 'no - v AB katalogu' a zaroven neni v letosnim 'Realnem' AB katalogu.
    if (include_internal.upper() == "YES" or include_internal.upper() == "NO - V AB KATALOGU") and in_real_ab_cat_2023 == "N/A":
        databaze_itemu[item]["TO INCLUDE 'Realny' INTERNAL 2023"] = 'yes'
    else:
        databaze_itemu[item]["TO INCLUDE 'Realny' INTERNAL 2023"] = 'no'    

print(f'INCLUDE "REALNY" INTERNAL KATALOG OK 17')






## 
## # 11.3 Stanoveni BFE ceny 2023 na zaklade vyse BFE eskalace.
## zahlavi.append("Vseobecna BFE price 2023 [EUR]")
## zahlavi.append("Vseobecna BFE price 2023 [USD]")
## 
## man_bfe_placard_2023_price_eur = 57
## man_bfe_id_placard_2023_price_eur = 126.03
## 
## 
## for item in databaze_itemu:
##     price_2023_eur = "?"    
##     # print(item)
##     # Cenu urcovat jen pro polozky, ktere pujdou do 2023 katalogu.   
##     if databaze_itemu.get(item).get("To include in 2023 BFE catalogue (sales history, low margin, comment, actual sales price not POR/N/A, correct SG)").upper() == "YES":
##         # Ziskani kurzu USD/EUR na zaklade, kdy byla polozka historicky nacenena.
##         usd_eur_ratio = float(databaze_itemu.get(item).get("prepocet USD/EUR podle roku naceneni").replace(",","."))
##                
##         # Pokud se jedna o MAN PLACARDY nebo ID PLACARDY → jednotnou cenu za placard.
##         man_placard = je_to_man_placard(item, databaze_itemu)
##         id_placard = je_to_id_placard(item, databaze_itemu)
##         
##         if man_placard:
##             if id_placard:
##                 databaze_itemu[item]["Vseobecna BFE price 2023 [EUR]"] = str(man_bfe_id_placard_2023_price_eur).replace(".",",")
##                 databaze_itemu[item]["Vseobecna BFE price 2023 [USD]"] = str(man_bfe_id_placard_2023_price_eur * eur_usd).replace(".",",")
##             else:
##                 databaze_itemu[item]["Vseobecna BFE price 2023 [EUR]"] = str(man_bfe_placard_2023_price_eur).replace(".",",")
##                 databaze_itemu[item]["Vseobecna BFE price 2023 [USD]"] = str(man_bfe_placard_2023_price_eur * eur_usd).replace(".",",")   
##         else:    
## 
##             # Pokud se jedna o polozky, ktere jsou i v SFE AB katalogu pro 2023 (MIX), dat stejnou USD a EUR cenu jako maji v SFE katalogu (chceme aby cena byla stejna).
##             if databaze_itemu.get(item).get("SFE 2023 katalog poslano do Francie 100 % cena [USD]") != "N/A": # Kontrola, zda je item i v SFE katalogu (nema N/A).
##                 # Rovnou se da USD cena jako kopie udaje z SFE katalogu.
##                 sfe_standard_price_usd = float(databaze_itemu.get(item).get("SFE 2023 katalog poslano do Francie 100 % cena [USD]").replace(",","."))
##                 # Eur cena se dopocita pomoci usd/eur kurzu.
##                 eur_price_calculated = sfe_standard_price_usd / usd_eur_ratio
## 
##                 databaze_itemu[item]["Vseobecna BFE price 2023 [EUR]"] = str(eur_price_calculated).replace(".",",")
##                 databaze_itemu[item]["Vseobecna BFE price 2023 [USD]"] = str(sfe_standard_price_usd).replace(".",",")
##                 
##             
##             # Jinak se cena spocita klasicky pro BFE polozku.
##             else:
##                 # print(item)
##                 # Ziskani EUR ceny 2022
##                 price_2022_eur = float(databaze_itemu.get(item).get("Actual Sales price now [EUR]").replace(",","."))
## 
##                 # Ziskani USD ceny 2022
##                 # Pokud neni cena v historicalu, mela by byt v prubeznem naceneni.
##                 if "N/A" in databaze_itemu.get(item).get("BFE historical 2022 [USD]"):
##                     if "N/A" in databaze_itemu.get(item).get("BFE pr. Nac 2022"):
##                         price_2022_usd = price_2022_eur * usd_eur_ratio 
##                     else:
##                         price_2022_usd = float(databaze_itemu.get(item).get("BFE pr. Nac 2022").replace(",",".")) * usd_eur_ratio 
##                 else:            
##                     price_2022_usd = float(databaze_itemu.get(item).get("BFE historical 2022 [USD]").replace(",","."))
## 
##                 # 2023 EUR cena pocitana jako EUR 2022 cena * (1 + eskalacni procento).
##                 price_2023_eur = round(price_2022_eur * eskalacni_faktor_bfe_vseobecne, 2)
##                 # 2023 USD cena pocitana jako 2022 USD cena * (1 + eskalacni procento).
##                 price_2023_usd = round(price_2022_usd * eskalacni_faktor_bfe_vseobecne, 2)
##                 
##                 databaze_itemu[item]["Vseobecna BFE price 2023 [EUR]"] = str(price_2023_eur).replace(".",",")
##                 databaze_itemu[item]["Vseobecna BFE price 2023 [USD]"] = str(price_2023_usd).replace(".",",")
##     else:
##         databaze_itemu[item]["Vseobecna BFE price 2023 [EUR]"] = "POR"
##         databaze_itemu[item]["Vseobecna BFE price 2023 [USD]"] = "POR"




output_to_write = list()
output_to_write.append(zahlavi)

for item in databaze_itemu:    
    line_to_to_write = list()    
    for key in databaze_itemu.get(item):
        line_to_to_write.append(databaze_itemu[item][key])
    output_to_write.append(line_to_to_write)

with open("SFE output.txt", "w", encoding='utf-8') as output:
    for line in output_to_write:    
        to_write = "\t".join(line)
        output.write(f'{to_write}\n')
    output.close()