import sys
import datetime

### FUNKCE

def struktura_kusovniku(kusovnik, hloubka_kusovniku, result_to_print): # VYTISKNE STRUKTURU KUSOVNIKU S JEHO LEVELY NA ZAKLADE DICT STRUKTRURY KUSOVNIKU.
 
    # print(f'Hloubka pred spustenim = {hloubka_kusovniku}')
    hloubka_kusovniku += 1
    # print(f'Hloubka pred spustenim = {hloubka_kusovniku}')
    for dil, qty_a_poddily in kusovnik.items(): 
        # print(f'Dil: {dil}')        
        # print(f'Dil bez projektu: {dil_bez_projektu}')
        # print(f'{"--"* hloubka_kusovniku}Level_{hloubka_kusovniku}: {dil}')
        print(f'{" "*hloubka_kusovniku}Level_{hloubka_kusovniku}: {dil}, BOM Qty: {kusovnik[dil]["qty"]}')
        result_to_print.append(f'Level_{hloubka_kusovniku}: {dil}, BOM Qty: {kusovnik[dil]["qty"]}')
        # print(result_to_print)
        for key, values in qty_a_poddily.items():
            if key == "poddily":
                struktura_kusovniku(values, hloubka_kusovniku, result_to_print)

def data_headings(data): # Vytvoreni zahlavi sloupcu z reportu.
    for line in data:
        if "***" in line[2:5]: # Identifikator linky se zahlavim z cq reportu.
            data_headings = [pole.strip() for pole in line.split("|")]
            data_headings.remove("***")
            break
    return data_headings

def import_data_cleaning(data): # Ocisteni dat a nahrazeni prazdnych poli za "0".
    cl_data = list()    
    for line in data: # Ocisteni dat a 
        if line[0] == "|":
            linka = [pole.strip() for pole in line.split("|")] # Ocisteni dat.
            for i, pole in enumerate(linka): # nahrazeni praydnych poli za "0".
                if len(pole) == 0:
                    linka[i] = "0"
                # smazani prvniho pole(vznikne s hodnotou 0 pri rozdeleni linky).
            del linka[0]
            cl_data.append(linka)
    return cl_data

def data_date_formating(data, data_headings): # Prevedeni pole datumu na date format.
    datumy = [data_headings.index(heading) for heading in data_headings if "EXDATE" in heading.upper()]
    # print(f' datumy indexy : {datumy}')
    for line in data: # Prevedeni pole datumu na date format.
        # print(line)
        # print(datumy_indexy)
        for datum_index in datumy:      
            
            if line[datum_index] != "0": # Preskoci prazdna pole (0)
                den, mesic, rok = line[datum_index].split("/")
                # print(rok, mesic, den)
                datum = datetime.date(int(rok), int(mesic), int(den))
                line[datum_index] = datum
    return data

def generic_database(data, headings): # Vytvoreni dictionary databaze jednotlivych itemu s jejich linkamy a daty z reportu.
    database_dictionary = dict() # Vsechny vrcholy z reportu a jejich linky priprava dict.
    
    for heading in headings:
        if heading.upper() == "ITEM":
            item_index = headings.index(heading) 

    for line in data:  
        # print(line)
        item = line[item_index]
        # print(f'Item: {item}')
        compile_line_dict = {} # Jednotlive linky vrcholu z dat
        data_dict = {} # Samotna data na linkach.
        # Sestaveni dat linky pro databazi.
        for i, data_field in enumerate(line):
            data_dict[headings[i]] = data_field # Sestaveni dat z linky jako dict s jmeny sloupcu zahlavi reportu jako klic.
        # print(f'Item: {item}')
        if item not in database_dictionary:# Pokud item jeste neni v databazi → vytvori ho tam s touto prvni linkou.
            # print(f'Item: {item} pridan do databaze')
            database_dictionary[item] = data_dict # Sestaveni linek predchoziho vrcholu jako dict vrchol : jeho linky s daty viz. vyse.
            # print(war_loc_database_dictionary)
            # print(f'databaze: {database_dictionary}')
    return database_dictionary

### PROGRAM

## 1. Sestaveni kusovniku po linkach z exportu BOM z CQ (vcetne expired linek)

# 1.a Nacteni dat z cq reportu BOMu. 

with open("Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\8_AFT_boudy_nacenovani\\CQ kusovnik export\\601567-20001+E01+Chybejici dily z VSB.txt", "r", encoding="Windows-1250") as bom_export:
    bom_data = bom_export.readlines()
    bom_export.close()

# 1.b Sestaveni zahlavi sloupcu dat z cq reportu.
bom_zahlavi = data_headings(bom_data)
print(f'Pocet sloupcu zahlavi: {len(bom_zahlavi)}')
# for nazev in bom_zahlavi:
#     # Identifikace sloupcu s Item pn.
#     if "ITEM" in nazev.upper() and "TYPE" not in nazev.upper():
#         print(nazev) 

# 1.c Ocisteni dat z cq reportu BOMu.
bom_data = import_data_cleaning(bom_data)

# 1.d Opraveni formatu datumu v datech.
bom_data = data_date_formating(bom_data, bom_zahlavi)

print(f'\nData kusovniku z CQ reportu nactena ...\n')

# 1.e Nalezeni indexu Item pn a expiry date pro jednotlive levely aby se pak dalo z dat vytahat vsechny platne linky kusovniku a sestavit effective date kusovnik.
effective_kusovnik_list = []
# Pridani zahlavi do kusovniku.
zahlavi_kusovniku = []
for nazev in bom_zahlavi:
    # print(nazev)
    if "ITEM" in nazev.upper() and "TYPE" not in nazev.upper():
        zahlavi_kusovniku.append(nazev)
effective_kusovnik_list.append(zahlavi_kusovniku)
# print(f'KUSOVNIK SE ZAHLAVIM: {kusovnik}')

# Prochazeni linek BOMu, sestavovani platnych linek (expiry date) a sestaveni platneho kusvoniku obsahujiciho pouze platne linky ve formatu : Dil, QTY BOM.
for line in bom_data:
    # print(line)
    linka_kusovniku = []
    for nazev in bom_zahlavi:
        # print(f'NAZEV: {nazev}')
        # NALEZENI INDEXU ITEMU V LINKACH.
        if "ITEM" in nazev.upper() and "TYPE" not in nazev.upper():
            pn_index = bom_zahlavi.index(nazev)
            if line[pn_index] != "0":
                # print(f'PN: {line[pn_index]}')         
                exdate = "default"
                # NALEZENI EX. DATE PRO KAZDY ITEM V LINCE.
                for nazev in bom_zahlavi[pn_index:]:
                    # BERE SE PRVNI NALEZENE EXDATE VPRAVO OD ZKOUMANEHO ITEMU.
                    if "EXDATE" in nazev.upper():
                        exdate_index = pn_index+bom_zahlavi[pn_index:].index(nazev)
                        exdate = line[exdate_index]
                        # print(exdate)
                        break # Az se najde prvni exdate (k momentalnimu pn) → hned break → POKRACOVANI NA DALSI ITEM V LINCE.         
                item_bom_qty = "default"
                # NALEZENI BOM QTY PRO KAZDY ITEM V LINCE.
                for nazev in bom_zahlavi[pn_index:]:
                    # BERE SE PRVNI NALEZENE BOM QTY VPRAVO OD ZKOUMANEHO ITEMU.
                    if "BOM QTY" in nazev.upper():
                        bom_qty_index = pn_index+bom_zahlavi[pn_index:].index(nazev)
                        bom_qty = line[bom_qty_index]
                        # print(bom_qty)
                        break # Az se najde prvni bom_qty (k momentalnimu pn) → hned break → POKRACOVANI NA DALSI ITEM V LINCE.                              
                # print(pn_index)
                # SESTVENI PLATNE LINKZ POUZE Z PLATNYCH ITEMU.
                if exdate > datetime.date.today():
                    ## dil_dict = dict()
                    ## qty_dict = dict()
                    ## qty_dict["QTY"] = bom_qty
                    
                    ## dil_dict[line[pn_index]] = qty_dict                    
                    dil = line[pn_index]

                    # print(f'DIL: {dil}')
                    # print(f'QTY: {bom_qty}')
                    # print(f'dil, QTY: {dil, bom_qty}')
                    linka_kusovniku.append([dil, bom_qty])
                    # print(linka_kusovniku)
                else:
                    break # Pokud ma dil v lince kusovniku neplatne extate → jeho poddily az nepridavat.
    if linka_kusovniku not in effective_kusovnik_list:
        effective_kusovnik_list.append(linka_kusovniku)        
print(f'\nEFFECTIVE KUSOVNIK (exdate linky ocistene) po linkach vytvoren ...\n')

# 2. Nacteni ITEMS DAT s informaci o ceně dílu.

with open('Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\8_AFT_boudy_nacenovani\\Item data\\items_data.txt', 'r', encoding= 'UTF-8') as idata:
    items_data = idata.readlines()

items_data_zahlavi = data_headings(items_data)
# print(items_data_zahlavi)
items_data = import_data_cleaning(items_data)
# print(items_data)
# for item in items_data:
    # print(item)

items_data_databaze = generic_database(items_data, items_data_zahlavi)
# print(items_data_databaze)
# for item, data in items_data_databaze.items():
    # print(f'{item}: {data}')

print(f'\nITEMS data polozek nactena ...\n')

# 3. Nacteni CMM dat PRO POROVNANI DILU, ZDA JSOU V CMM A BUEME CHTIT PRIDAVAT DO BOUDY.
# with open('C:\\Users\\Ondrej.rott\\Documents\AFT nove boudy\\23_Python_utilities\\1_LT + Nastavení položek\\CMM 601567-20001.txt', 'r', encoding= 'UTF-8') as file:
with open('Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\8_AFT_boudy_nacenovani\\CMM data\\CMM 601567-20001.txt', 'r', encoding= 'UTF-8') as file:
    cmm_data = file.readlines()

cmm_zahlavi = data_headings(cmm_data)
# print(cmm_zahlavi)
cmm_data = import_data_cleaning(cmm_data)
# print(cmm_data)
cmm_databaze = generic_database(cmm_data, cmm_zahlavi)
# print(cmm_databaze)

print(f'\nCMM data polozek nactena ...\n')

# 4. Prochazeni linek kusovniku listu a overovani, jestli dily jsou v CMM a jestli jejich naddily jsou v CMM.

dily_v_cmm = set()
# Set dily, pod kterymi je nekde dil z CMM (nemusi nutne byt primo poddil.)
dily_maji_poddil_v_cmm = set()

for linka in effective_kusovnik_list[1:]:
    for dil in linka:
        pn = dil[0]# effective kusovnik je ve formatu EK = [[pn1, qty],[pn2,qty],...] 
        if pn in cmm_databaze:
            if pn not in dily_v_cmm:            
                # print(f'Dil {dil} je v CMM databazi')
                dily_v_cmm.add(pn)
            for naddil in linka[0:linka.index(dil)]:
                nadpn = naddil[0]
                if nadpn not in dily_maji_poddil_v_cmm: 
                    dily_maji_poddil_v_cmm.add(nadpn)
                    # print(f'Naddil {naddil} dilu {dil} pridan do seznamu dilu, ktere maji poddil v cmm.')     
print(f'Seznam unikatnich dilu z kusovniku, ktere jsou v CMM vztvoren ...\n')
print(f'{len(dily_v_cmm)} unikatnich dilu z kusovniku jsou v CMM databazi.\n')
# print(dily_v_cmm)
# for dil in dily_v_cmm:
#     print(dil)

print(f'\nSeznam unikatnich dilu z kusovniku, ktere maji nejaky poddil v CMM vytvoren ...\n')
print(f'{len(dily_maji_poddil_v_cmm)} unikatnich dilu z kusovniku ma poddil v CMM databazi.')
# for dil in dily_maji_poddil_v_cmm:
#     print(dil)

# 5. Nalezeni duplicitnich CMM dilu v efektivnim kusovniku. (Je potreba pro dalsi krok. 6)
kolikrat_je_item_v_effektovnim_kusovniku = dict()
for linka in effective_kusovnik_list:
    for dil in linka:
        pn = dil[0] # [pn, bom qty], ....
        if pn not in kolikrat_je_item_v_effektovnim_kusovniku:
            kolikrat_je_item_v_effektovnim_kusovniku[pn] = 1
        else:
            kolikrat_je_item_v_effektovnim_kusovniku[pn] += 1
print(f'Pocty jednotlivych dilu v effective kusovniku:')
# for dil, pocet in kolikrat_je_item_v_effektovnim_kusovniku.items():
#     print(dil, pocet)    
print(f'\nSeznam duplicitnich CMM dilu vytvoren ...')

# 6. Sestaveni nove verze kusovniku linkek pouze s itemy z CMM / itemy, ktere pod sebou maji CMM dily. Doplnujicne se overuje, "Petrovo pravidlo": Pokud dil neni v CMM a ma poze poddily, ktere jsou v CMM a jsou to vsechno duplikaty v boude → dil ani poddily nepridavat.
kusovnik_cmm_list = list()

for linka in effective_kusovnik_list[1:]:
    if len(linka) == 1:
        continue
    cmm_linka = list()
    predposledni_dil_linky = linka[-2]
    posledni_dil_linky = linka[-1]
    for dil in linka:
        pn = dil[0]
        #### nove pridany kod
        # Pokud poddil neni v CMM a uz v lince za sebou ma jen 1 vec a ta vec je CMM duplikat → nepidavat celou linku.
        if not(pn not in dily_v_cmm and predposledni_dil_linky[0] == pn and posledni_dil_linky[0] in dily_v_cmm and kolikrat_je_item_v_effektovnim_kusovniku.get(posledni_dil_linky[0]) > 1):
        ####        
        # Test, zda je to CMM dil, nebo ma pod sebou CMM dily:     
            if pn in dily_v_cmm or pn in dily_maji_poddil_v_cmm:
                cmm_linka.append(dil)
        else: # Pokuk dil nesplnuje kriteria vyse, dily za nim uz neproverovat → skocit na dalsi linku.
            break
    # Pokud poslednim dilem linky neni dil z CMM → linku nepridavat.
    # print(cmm_linka)
    if len(cmm_linka) != 0 and cmm_linka[-1][0] in dily_v_cmm:    
        if cmm_linka not in kusovnik_cmm_list:
            kusovnik_cmm_list.append(cmm_linka)

print(f'\nVysledny Upraveny kusovnik po linkach vytvoren ... ')
# print(f'CMM kusovnik: ')
# for line in kusovnik_cmm_list:
#     print(line)

# 6. Sestaveni struktury kusovniku CMM upraveneho.
kusovnik_cmm_databaze = dict()

for linka in kusovnik_cmm_list:
    # print(f'LINKA: {linka}')
    for level, item in enumerate(linka):
        pn = item[0]
        qty = item[1]
        # Vrchol
        if level == 0:
            if pn not in kusovnik_cmm_databaze:
                kusovnik_cmm_databaze[pn] = {"qty":qty, "poddily":dict()}
                # print(kusovnik_cmm_databaze)
        elif level == 1:           
            if pn not in kusovnik_cmm_databaze.get(linka[0][0]).get("poddily"):
                kusovnik_cmm_databaze[linka[0][0]]["poddily"][pn] = {"qty":qty, "poddily":dict()}
                # print(kusovnik_cmm_databaze)
        elif level == 2:
            if pn not in kusovnik_cmm_databaze.get(linka[0][0]).get("poddily").get(linka[1][0]).get("poddily"):
                kusovnik_cmm_databaze[linka[0][0]]["poddily"][linka[1][0]]["poddily"][pn] = {"qty":qty, "poddily":dict()}
                # print(kusovnik_cmm_databaze)           
        elif level == 3:
            if pn not in kusovnik_cmm_databaze.get(linka[0][0]).get("poddily").get(linka[1][0]).get("poddily").get(linka[2][0]).get("poddily"):
                kusovnik_cmm_databaze[linka[0][0]]["poddily"][linka[1][0]]["poddily"][linka[2][0]]["poddily"][pn] = {"qty":qty, "poddily":dict()}
                # print(kusovnik_cmm_databaze) 
        elif level == 4:
            if pn not in kusovnik_cmm_databaze.get(linka[0][0]).get("poddily").get(linka[1][0]).get("poddily").get(linka[2][0]).get("poddily").get(linka[3][0]).get("poddily"):
                kusovnik_cmm_databaze[linka[0][0]]["poddily"][linka[1][0]]["poddily"][linka[2][0]]["poddily"][linka[3][0]]["poddily"][pn] = {"qty":qty, "poddily":dict()}
                # print(kusovnik_cmm_databaze)
        elif level == 5:
            if pn not in kusovnik_cmm_databaze.get(linka[0][0]).get("poddily").get(linka[1][0]).get("poddily").get(linka[2][0]).get("poddily").get(linka[3][0]).get("poddily").get(linka[4][0]).get("poddily"):
                kusovnik_cmm_databaze[linka[0][0]]["poddily"][linka[1][0]]["poddily"][linka[2][0]]["poddily"][linka[3][0]]["poddily"][linka[4][0]]["poddily"][pn] = {"qty":qty, "poddily":dict()}
                # print(kusovnik_cmm_databaze)
        elif level == 6:
            if pn not in kusovnik_cmm_databaze.get(linka[0][0]).get("poddily").get(linka[1][0]).get("poddily").get(linka[2][0]).get("poddily").get(linka[3][0]).get("poddily").get(linka[4][0]).get("poddily").get(linka[5][0]).get("poddily"):
                kusovnik_cmm_databaze[linka[0][0]]["poddily"][linka[1][0]]["poddily"][linka[2][0]]["poddily"][linka[3][0]]["poddily"][linka[4][0]]["poddily"][linka[5][0]]["poddily"][pn] = {"qty":qty, "poddily":dict()}
                # print(kusovnik_cmm_databaze)
        elif level == 7:
            if pn not in kusovnik_cmm_databaze.get(linka[0][0]).get("poddily").get(linka[1][0]).get("poddily").get(linka[2][0]).get("poddily").get(linka[3][0]).get("poddily").get(linka[4][0]).get("poddily").get(linka[5][0]).get("poddily").get(linka[6][0]).get("poddily"):
                kusovnik_cmm_databaze[linka[0][0]]["poddily"][linka[1][0]]["poddily"][linka[2][0]]["poddily"][linka[3][0]]["poddily"][linka[4][0]]["poddily"][linka[5][0]]["poddily"][linka[6][0]]["poddily"][pn] = {"qty":qty, "poddily":dict()}
                # print(kusovnik_cmm_databaze)
        elif level == 8:
            if pn not in kusovnik_cmm_databaze.get(linka[0][0]).get("poddily").get(linka[1][0]).get("poddily").get(linka[2][0]).get("poddily").get(linka[3][0]).get("poddily").get(linka[4][0]).get("poddily").get(linka[5][0]).get("poddily").get(linka[6][0]).get("poddily").get(linka[7][0]).get("poddily"):
                kusovnik_cmm_databaze[linka[0][0]]["poddily"][linka[1][0]]["poddily"][linka[2][0]]["poddily"][linka[3][0]]["poddily"][linka[4][0]]["poddily"][linka[5][0]]["poddily"][linka[6][0]]["poddily"][linka[7][0]]["poddily"][pn] = {"qty":qty, "poddily":dict()}
                # print(kusovnik_cmm_databaze)
        elif level == 9:
            if pn not in kusovnik_cmm_databaze.get(linka[0][0]).get("poddily").get(linka[1][0]).get("poddily").get(linka[2][0]).get("poddily").get(linka[3][0]).get("poddily").get(linka[4][0]).get("poddily").get(linka[5][0]).get("poddily").get(linka[6][0]).get("poddily").get(linka[7][0]).get("poddily").get(linka[8][0]).get("poddily"):
                kusovnik_cmm_databaze[linka[0][0]]["poddily"][linka[1][0]]["poddily"][linka[2][0]]["poddily"][linka[3][0]]["poddily"][linka[4][0]]["poddily"][linka[5][0]]["poddily"][linka[6][0]]["poddily"][linka[7][0]]["poddily"][linka[8][0]]["poddily"][pn] = {"qty":qty, "poddily":dict()}
                # print(kusovnik_cmm_databaze)
        elif level == 10:
            if pn not in kusovnik_cmm_databaze.get(linka[0][0]).get("poddily").get(linka[1][0]).get("poddily").get(linka[2][0]).get("poddily").get(linka[3][0]).get("poddily").get(linka[4][0]).get("poddily").get(linka[5][0]).get("poddily").get(linka[6][0]).get("poddily").get(linka[7][0]).get("poddily").get(linka[8][0]).get("poddily").get(linka[9][0]).get("poddily"):
                kusovnik_cmm_databaze[linka[0][0]]["poddily"][linka[1][0]]["poddily"][linka[2][0]]["poddily"][linka[3][0]]["poddily"][linka[4][0]]["poddily"][linka[5][0]]["poddily"][linka[6][0]]["poddily"][linka[7][0]]["poddily"][linka[8][0]]["poddily"][linka[9][0]]["poddily"][pn] = {"qty":qty, "poddily":dict()}
                # print(kusovnik_cmm_databaze)
# print(kusovnik_cmm_databaze)

# 7. Tisk struktury kusovniku upraveneho tak aby zahrnoval pouze dily, ktere chceme dat do boudy.
print(f'\nVysledna podoba upraveneho kusovniku, obsahujiciho pouze dily, ktere chceme nacenit do AFT boudy jako nahradni dily.\n')


global result_to_print
result_to_print = list() 
struktura_kusovniku(kusovnik_cmm_databaze, -1, result_to_print)


with open("output.txt", "w", encoding= "UTF-8") as output:
    output.write(f'Struktura vysledenho kusovniku AFT boudy {effective_kusovnik_list[1:][0][0][0]} k naceni\n\n')
    for line in result_to_print:
        output.write(f'{line}\n')


# 9. Tisk struktury kusovniku upraveneho tak aby zahrnoval pouze dily, ktere chceme dat do boudy + jejich sales cenu USD.
print(f'\nVysledna podoba upraveneho kusovniku, obsahujiciho pouze dily, ktere chceme nacenit do AFT boudy jako nahradni dily a jejich cena USD.\n')
