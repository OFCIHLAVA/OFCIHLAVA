import datetime

def data_import(cq_export_cesta): # Nacteni dat z CQ reportu.
    with open(cq_export_cesta,'r', encoding='Windows-1250') as input_file:
        data = input_file.readlines()
        input_file.close()
    return data

def data_headings(data): # Vytvoreni zahlavi sloupcu z reportu.
    for line in data:
        if "***" in line[2:5]: # Identifikator linky se zahlavim z cq reportu
            data_headings = [pole.strip() for pole in line.split("|")]
            del data_headings[0]
            break
    return data_headings  

def import_data_cleaning(data): # Ocisteni dat a nahrazeni prazdnych poli za "0".
    cl_data = list()    
    for line in data: # Ocisteni dat a 
        if line[0] == "|":
            linka = [pole.strip() for pole in line.split("|")] # Ocisteni dat.
            # print(linka)
            for i, pole in enumerate(linka): # nahrazeni praydnych poli za "0".
                if len(pole) == 0:
                    linka[i] = "0"
            # smazani prvniho pole(vznikne s hodnotou 0 pri rozdeleni linky).
            del linka[0]
            cl_data.append(linka)
            data = cl_data
    return data

def data_date_formating(data, data_headings): # Prevedeni pole datumu na date format.
    datumy = [data_headings.index(heading) for heading in data_headings if "DATE" in heading.upper()]

    for line in data: # Prevedeni pole datumu na date format.
        
        # print(datumy_indexy)
        for datum_index in datumy:      
            den, mesic, rok = line[datum_index].split("/")
            # print(rok, mesic, den)
            datum = datetime.date(int(rok), int(mesic), int(den))
            line[datum_index] = datum
    return data

def generic_database(data, headings, primary_key): # Vytvoreni dictionary databaze jednotlivych itemu s jejich linkamy a daty z reportu.
    database_dictionary = dict() # Vsechny vrcholy z reportu a jejich linky priprava dict.
    
    for heading in headings:
        if heading.upper() == str(primary_key).upper():
            pk_index = headings.index(heading) 

    for line in data:  
        key = line[pk_index]
        compile_line_dict = {} # Jednotlive linky vrcholu z dat
        data_dict = {} # Samotna data na linkach.
        # Sestaveni dat linky pro databazi.
        for i, data_field in enumerate(line):
            data_dict[headings[i]] = data_field # Sestaveni dat z linky jako dict s jmeny sloupcu zahlavi reportu jako klic.
        # print(f'Item: {item}')
        if key not in database_dictionary:# Pokud item jeste neni v databazi → vytvori ho tam s touto prvni linkou.
            compile_line_dict[1] = data_dict # Sestaveni jednotlivych linek jako dict cislo linky vrcholu : data v lince vyse.
            # print(compile_line_dict)             
            # print(f'Neni tam zatim.')
            database_dictionary[key] = compile_line_dict # Sestaveni linek predchoziho vrcholu jako dict vrchol : jeho linky s daty viz. vyse.
            # print(war_loc_database_dictionary)
        else:# Pokud item uz je v databazi..
            # print(f'Uz tam je.')
            # Pokud pridavana lokace tam jeste neni → prodat samostatnou linku.
            database_dictionary[key][len(database_dictionary[key])+1] = data_dict
            # print(war_loc_database_dictionary)
    return database_dictionary


# Import dat:
data = data_import("Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\10_Zavírání projektu\\projekty report\\projekty.txt")

# Zahlavi dat
zahlavi = data_headings(data)

# Ocisteni dat
data = import_data_cleaning(data)

# Opraveni formatu datumu
data = data_date_formating(data, zahlavi)

# # Kontrolni tisk dat
# print(zahlavi)
# for line in data:
#     print(line)

# Vytvoreni databaze Sales order linek
so_databaze = generic_database(data, zahlavi, "Sales Order")

# for so, linky in so_databaze.items():
#     print(so)
#     for linka, data in linky.items():
#         print(linka, data)

# Vybrani projektu k uzavreni + podezrelych SO
# Projit vsechny linky SO a kouknout se, jestli je uz invoicnuto. Pokud ano → pridat projekt linky k seznamu uzaviranych projektu linky. Return seznam projektu k uzavreni.
# Pokud nektera linka neni jeste invoicnuta → overit, jestli je ordered Qty na ni <= 0 , pokud je → Vratit error hlaseni k prochazne SO s touto linkou / linkami k dalsimu provereni.
#  Pokud vsechny ordered qty >0 → vsechnt projekty ze SO brat jako jeste neuzavirane a preskocit na dalsi SO.



print(f'\nTeď budu prověřovat všechny Sales order linky všech Aktivních projektů, jestli už jsou invoicnuté. Pokud ano, projekty z dané Sales ordery půjdou k uzavření . . .\n')

# Globalni trackovani stavu linek 
projekty_k_zavreni = set()
projekty_nezavirat = set()

zatim_neinvoicnute_linky = set()
divne_so_proverit = set()

for so, linky in so_databaze.items():
    
    # Tracker stavu linek pro current SO.
    so_projekty_k_zavreni = set()
    
    
    
    so_zatim_neinvoicnute_linky = set()
    so_divne_proverit = set()   
    
    print(f'\nProveruji SO: {so}...\n')
    for linka, data_linky in linky.items():
        print(f'    linka: {linka}')
        projekt = data_linky.get("Projekt")
        invoice_date = data_linky.get("Invoice Date")
        ordered_qty = float(data_linky.get("Ordered Qty"))
        
        # Pokud v SO nejaka neinvoicnuta linka:
        if invoice_date == datetime.date(1970, 1, 1):
            # Pokud se jedna o vratku / 0 QTY linku:
            if ordered_qty <= 0:
                print(f'\nPOZOR! sales order {so} ma v sobe podivnou linku s Qty ordered <= 0 {linka}: {data_linky}.Potreba proverit.')
                so_divne_proverit.add(str(data_linky).replace("\'",""))
                projekty_nezavirat.add(projekt)
                if projekt in projekty_k_zavreni:
                    projekty_k_zavreni.discard(projekt)     
            else:
                print(f'\nSales order {so} ma v sobe zatim neinvocnutou linku {linka}: {data_linky}\nProjekty z teto SO se zatim nebudou uzavirat')
                so_zatim_neinvoicnute_linky.add(str(data_linky).replace("\'",""))
                projekty_nezavirat.add(projekt)
                if projekt in projekty_k_zavreni:
                    projekty_k_zavreni.discard(projekt)   
        else:
            if projekt not in projekty_nezavirat:
                so_projekty_k_zavreni.add(projekt)
    # Pokud SO ma pouze projektove linky uz invoicnute → vsechny projekty z ni uzavrit.
    if len(so_divne_proverit) == 0 and len(so_zatim_neinvoicnute_linky) == 0:
        print(f'\nSales order {so} ma uz vsechny linky s projektovymi dily invoicnute → Uzavrit projekty: {so_projekty_k_zavreni}')
        for projekt in so_projekty_k_zavreni:
            projekty_k_zavreni.add(projekt)
    else:
        for divna_so_linka in so_divne_proverit:
            divne_so_proverit.add(divna_so_linka)
        for neinvoicnuta_linka in so_zatim_neinvoicnute_linky:
            zatim_neinvoicnute_linky.add(neinvoicnuta_linka)

# Print Result + ulozeni do souboru output.
with open("Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\10_Zavírání projektu\\output_projekty.txt", "w", encoding = "UTF-8") as o:
    print('\nRESULT:\n')
    o.write('RESULT:\n')

    print('\nPROJEKTY uzavrit:\n')
    o.write('\nPROJEKTY uzavrit:\n')
    for projekt in projekty_k_zavreni:
        print(projekt)
        o.write(f'{projekt}\n')

    print('\nZatim neinvoicnute linky:\n')
    o.write('\nZatim neinvoicnute linky:\n\n')
    for linka in zatim_neinvoicnute_linky:
        print(linka)
        o.write(f'{linka}\n')

    print('\nDVINE linky (proverit):\n')
    o.write('\nDVINE linky (proverit):\n\n')
    for linka in divne_so_proverit:        
        print(linka)
        o.write(f'{linka}\n')
    o.close()

print(f'\nHOTOVO\n')
print(f'\nVýsledek byl uložen do souboru output.txt\n')
input('\nPROGRAM UKONCI stiskem ENTER ...')