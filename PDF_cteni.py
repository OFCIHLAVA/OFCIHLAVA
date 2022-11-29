from unicodedata import digit

# Nacte txt soubor se zkopirovanymi daty z PDF souboru jako list linek.
with open("Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\7_Vyhledavac PN v PDF souborech\\data.txt","r", encoding='utf-8') as file:
    data = file.readlines()
    file.close()
# print(data)
#print(len(data))

print(f'\nMEZERA\n')

upravena_data = []
seznam_potencialnich_pn = []

# Ocisteni dat a rozdeleni textu podle mezer " ".
for radek in data:
    radek = radek.strip()
    radek = radek.replace("\n","")
    radek = radek.split(" ")
    upravena_data.append(radek)
# print(upravena_data)

for radek in upravena_data:
    print(radek)

# # Zapise linky seznam do txt souboru "o_linky.txt"
# with open("Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\7_Vyhledavac PN v PDF souborech\\o_linky.txt", "w", encoding='utf-8') as ol:
#     ol.write("Seznam linek:\n\n")
#     ol.write(f'Pocet linek: {len(upravena_data)}.\n\n')
#     for radek in upravena_data:
#         ol.write(f'{radek}\n')
#     ol.write("\nKonec seznamu")
#     ol.close()


print(f'\nMEZERA1\n')

# Projde postupne vsechny textove retezce kazde linky a na zaklade podminek nize vyhodnoti, zda se potencialne muze jednat o platny P/N number. Pokud ano, prida ho do seznamu vysledku.
for radek in upravena_data:
    for text in radek:
        if len(text) > 5: # zadne PN v LNku jsem nenasel ze by bylo kratsi nez 6. Pokud ano, je nutno upravit podminku.
            if not text.isalpha() : # Pokud se nejedna ciste o text.
                # kontrola jestli ma v sobe alespon 1 cislici 0-9.
                digits_count = 0
                for char in text:
                    if char.isdigit():
                        digits_count +=1
                if digits_count != 0 and text not in seznam_potencialnich_pn:
                    if text[-1] != "-":    
                        seznam_potencialnich_pn.append(text)
print(seznam_potencialnich_pn)

# Zapise vznikly seznam do txt souboru "outpu.txt"
with open("Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\7_Vyhledavac PN v PDF souborech\\output.txt", "w", encoding='utf-8') as o:
    o.write("Seznam exportovanych dilu:\n\n")
    o.write(f'Pocet unikatni pn: {len(seznam_potencialnich_pn)}.\n\n')
    for pn in seznam_potencialnich_pn:
        o.write(f'{pn}\n')
    o.write("\nKonec seznamu")
    o.close()

input(f'\nPress ENTER to EXIT program ...')