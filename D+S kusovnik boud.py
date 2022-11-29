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

with open("Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\8_AFT_boudy_nacenovani\\CQ kusovnik export\\601253-201.txt", "r", encoding="Windows-1250") as bom_export:
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

# 5. Nalezeni duplicitnich CMM dilu v efektivnim kusovniku. (Je potreba pro dalsi krok. 6)
kolikrat_je_item_v_effektovnim_kusovniku = dict()
for linka in effective_kusovnik_list[1:]:
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


# 7. Tisk struktury kusovniku upraveneho tak aby zahrnoval pouze dily, ktere chceme dat do boudy.
print(f'\nVysledna podoba upraveneho kusovniku, obsahujiciho pouze dily, ktere chceme nacenit do AFT boudy jako nahradni dily.\n')

with open("AB_output.txt", "w", encoding= "UTF-8") as output:
    output.write(f'Unikatni dily v efektovnim kusovniku AB boudy {effective_kusovnik_list[1:][0][0][0]}\n\n')
    for dil in kolikrat_je_item_v_effektovnim_kusovniku:
        output.write(f'{dil}\n')
