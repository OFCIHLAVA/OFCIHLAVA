import funkce_prace

file ='C:\\Users\\Ondrej.rott\\Documents\\Python\\programy\\Exporty z LN\\report.txt'

slinka = funkce_prace.import_kusovniku_LN(file)
print(f'hotovo 1')
boudy = funkce_prace.seznam_boud_z_importu(slinka)
print(f'hotovo 2')
vysledek = funkce_prace.boudy_s_kusovnikem(boudy, slinka)
print(f'hotovo 3')
with open("output.txt", "w") as o:
    o.write(vysledek)
    o.close()
