import funkce_prace
import os

# Pridani novycn boud do reportu.

# file_to_process = "report.txt"
# path_to_export = 'C:\\Users\\Ondrej.rott\\Documents\\Python\\programy\\Exporty z LN\\' + str(file_to_process)

# ocistene=funkce_prace.import_kusovniku_LN(path_to_export)
# unikatni_boudy = funkce_prace.seznam_boud_z_importu(ocistene)
# output  = funkce_prace.boudy_s_kusovnikem(unikatni_boudy, ocistene)

# with open("output.txt", "w") as output_file:
  # output_file.write(output)
  # output_file.close


# Samotne dotazovani

path_kusovniky_databaze = 'C:\\Users\\Ondrej.rott\\Documents\\Python\\programy\\databaze boud s kusovniky.txt'
path_program_databaze = 'C:\\Users\\Ondrej.rott\\Documents\\Python\\programy\\seznam programu.txt'

databaze_pro_dotaz = funkce_prace.nacteni_databaze_boud_pro_dotaz(path_kusovniky_databaze)
print("Data nactena a pripravena pro dotazovani...")

kvp_programy = funkce_prace.programy_boud(path_program_databaze)
print("Databaze prgramu vsech boud nacetena...\n")

# Samotne dotazovani

multivysledek = []

while True:
    dotaz = input("Zadej PN, nebo seznam PN z bunek pod sebou v excelu, u kterych chces zjistit program / nakonec press ENTER pro vysledek:\n")
    if dotaz == "":
      break
    dotaz = dotaz.strip()
    vysledek = funkce_prace.dotaz_pn_program(dotaz, databaze_pro_dotaz, kvp_programy)
    print(str(vysledek))
    if vysledek not in multivysledek:
       multivysledek.append(vysledek)
    # print("\n")
for vysledek in multivysledek:
  print(vysledek)