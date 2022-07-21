from unicodedata import digit


with open("test.txt", "r") as file:
    data = file.readlines()
    file.close()
print(data)
#print(len(data))

print(f'\nMEZERA\n')

upravena_data = []
upravena_data1 = []

for radek in data:
    radek = radek.strip()
    radek = radek.replace("\n","")
    radek = radek.split(" ")
    upravena_data.append(radek)
print(upravena_data)

print(f'\nMEZERA1\n')

for radek in upravena_data:
    for text in radek:
        if len(text) > 5: # zadne PN v LNku jsem nenasel ze by bylo kratsi nez 6. Pokud ano, je nutno upravit podminku.
            if not text.isalpha() : # Pokud se nejedna ciste o text.
                # kontrola jestli ma v sobe alespon 1 cislici 0-9.
                digits_count = 0
                for char in text:
                    if char.isdigit():
                        digits_count +=1
                if digits_count != 0 and text not in upravena_data1:
                    upravena_data1.append(text)
print(upravena_data1)

with open("output.txt", "w") as o:
    o.write("Seznam exportovanych dilu:\n\n")
    o.write(f'Pocet unikatni pn: {len(upravena_data1)}.\n\n')
    for pn in upravena_data1:
        o.write(f'{pn}\n')
    o.write("\nKonec seznamu")
    o.close()