from moduly.funkce import funkce_prace
from logging import root
from tkinter import RAISED, Button, Label, StringVar, Tk, END, Text

# Samotne dotazovani

path_kusovniky_databaze = 'Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\5_SFExBFExMIX\\databaze\\databaze boud s kusovniky.txt'
path_program_databaze = 'Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\5_SFExBFExMIX\\databaze\\seznam programu.txt'

databaze_kusovniku_pro_dotaz = funkce_prace.nacteni_databaze_boud_pro_dotaz(path_kusovniky_databaze)
seznam_boud_z_databaze_kusovniku = [key for key in databaze_kusovniku_pro_dotaz]
print("Databaze bud s kusovniky nactena a pripravena pro dotazovani . . .")

kvp_programy_pro_dotaz = funkce_prace.programy_boud(path_program_databaze)
seznam_boud_z_databaze_programu = [key for key in kvp_programy_pro_dotaz]
print("Databaze programu vsech boud nacetenaa pripravena pro dotazovani . . .\n")

# Samotne dotazovani

root = Tk()

root.title("Ukazovadlo programů SFE x BFE x MIX")
root.iconbitmap("Y:\\Departments\\Sales and Marketing\\Aftersales\\11_PLANNING\\23_Python_utilities\\5_SFExBFExMIX\\graphics\\icon.ico")
root.geometry('900x600+0+3')

zadatPn = StringVar()
zadatPnLabel = Label(root,textvariable=zadatPn, font=('Calibry 10'))
zadatPn.set(f'↓ Tady zadej PN, nebo seznam PN z excelu pod sebou ↓')

outputLabel = Label(root,text=f'↓ Výsledné programy ↓', font=('Calibry 10'))
databaze_check_label = Label(root,text=f'↓ Jsou boudy v databázi? ↓', font=('Calibry 10'))
empty1 = Label(root, width = 15, height=1, text=f'', font=('Calibry 10'), highlightbackground="black", highlightcolor="black", highlightthickness=1, bd=3)

entry = Text(root, width = 30, height=8, borderwidth = 5, font=('Calibry 10'))
output = Text(root, width =40, height=8, borderwidth = 5,font=('Calibry 10'))
output['state'] = 'disabled'

entry_boudy = Text(root, width = 20, height=4, borderwidth = 5, font=('Calibry 10'))

output_boudy = Text(root,width=60, padx=50 ,borderwidth = 5,height=20, font=('Calibry 10'))
output_boudy['state'] = 'disabled'


def getProgram():
    output['state'] = 'normal'
    output.delete(1.0, END)
    dotaz = entry.get(1.0, END).split("\n")
    dotaz = [pn.strip() for pn in dotaz if pn != ""]    
    if len(dotaz) != 0:
        vysledek_dotazu = [funkce_prace.dotaz_pn_program(pn, databaze_kusovniku_pro_dotaz, kvp_programy_pro_dotaz) for pn in dotaz if pn != ""]
        for vysledek in vysledek_dotazu:
            output.insert(END, f'{vysledek[0]}\n')
    output['state'] = 'disabled'

def getBoudy():
    output_boudy['state'] = 'normal'
    output_boudy.delete(1.0, END)
    dotaz = entry.get(1.0, END).split("\n")
    dotaz = [pn.strip() for pn in dotaz if pn != ""]    
    if len(dotaz) != 0:
        vysledek_dotazu = [funkce_prace.dotaz_pn_program(pn, databaze_kusovniku_pro_dotaz, kvp_programy_pro_dotaz) for pn in dotaz if pn != ""]      
        print(vysledek_dotazu)
        for vysledek in vysledek_dotazu:
            if len(vysledek[1]) != 0: 
                poddil = vysledek[0].split(":")[0]
                obsazen_v_boudach = vysledek[1]

                output_boudy.insert(END, f'Dil {poddil} obsazen v boudach:\n\n')
                for bouda in obsazen_v_boudach:
                    output_boudy.insert(END, f'{poddil} : {bouda}\n')
                output_boudy.insert(END, f'\n')    
            else:
                output_boudy.insert(END, f'{vysledek[0].split(":")[0]} v boudach:\n[NELZE URCIT]\n\n')   
    output_boudy['state'] = 'disabled'

def dotaz_bouda_v_databazi():
    output_boudy['state'] = 'normal'

    output_boudy.delete(1.0, END)

    boudy_pro_dotaz = entry_boudy.get(1.0, END).split("\n")
    boudy_pro_dotaz = [bouda.replace("\n",".") for bouda in boudy_pro_dotaz if bouda != ""]
    print(boudy_pro_dotaz)

    boudy_z_dotazu_v_databazich_s_programem = set()

    for bouda in boudy_pro_dotaz:     
       v_programech = True if bouda in seznam_boud_z_databaze_programu else False
       v_kusovnicich = True if bouda in seznam_boud_z_databaze_kusovniku else False     
       program = kvp_programy_pro_dotaz.get(bouda, f'Neni v databazi programu.')
       bouda_s_programem = f'{bouda}:{program}:{v_programech}:{v_kusovnicich}'
       boudy_z_dotazu_v_databazich_s_programem.add(bouda_s_programem)

    print(boudy_z_dotazu_v_databazich_s_programem)

    output_boudy.insert(END, f'Bouda:program:je v databazi programu:je v databazi kusovniku\n')
    for bouda in boudy_z_dotazu_v_databazich_s_programem:
        output_boudy.insert(END, f'{bouda}\n')

    output_boudy['state'] = 'disabled'        

def obsah_cele_databaze():
    output_boudy['state'] = 'normal'
    output_boudy.delete(1.0, END)

    vsechny_output_boudy_v_databazich = set()
    vsechny_output_boudy_v_databazich_s_programem = set()
    
    for bouda in seznam_boud_z_databaze_kusovniku:
        if not " " in bouda:
            vsechny_output_boudy_v_databazich.add(bouda)

    for bouda in seznam_boud_z_databaze_programu:
        if not " " in bouda:
            vsechny_output_boudy_v_databazich.add(bouda)
    
    for bouda in vsechny_output_boudy_v_databazich:     
       v_programech = True if bouda in seznam_boud_z_databaze_programu else False
       v_kusovnicich = True if bouda in seznam_boud_z_databaze_kusovniku else False     
       program = kvp_programy_pro_dotaz.get(bouda, f'Neni v databazi programu.')
       bouda_s_programem = f'{bouda}:{program}:{v_programech}:{v_kusovnicich}'
       vsechny_output_boudy_v_databazich_s_programem.add(bouda_s_programem)

    output_boudy.insert(END, f'Bouda:program:je v databazi programu:je v databazi kusovniku\n')
    for bouda in vsechny_output_boudy_v_databazich_s_programem:
        output_boudy.insert(END, f'{bouda}\n')

    output_boudy['state'] = 'disabled'

def clear():
    entry.delete(1.0, END)
    entry_boudy.delete(1.0, END)
    
    output['state'] = 'normal'
    output.delete(1.0, END)
    output['state'] = 'disabled'

    output_boudy['state'] = 'normal'
    output_boudy.delete(1.0, END)
    output_boudy['state'] = 'disabled'

def copyToClipboardOutput():
    x = output.get(1.0, END)
    output.clipboard_clear()
    output.clipboard_append(x)

def copyToClipboardoutput_boudy():
    y = output_boudy.get(1.0, END)
    output_boudy.clipboard_clear()
    output_boudy.clipboard_append(y)

button_run = Button(root, text= f'Urcit program →', padx=20, pady=5, borderwidth=5, command=getProgram)
button_clear = Button(root, text= "CLEAR", padx=20, pady=5, borderwidth=5, command=clear)
button_boudy = Button(root, text= f'Obsazeno v boudach:', padx=20, pady=5, borderwidth=5, command=getBoudy)
button_copy_output = Button(root, text= f'Copy OUTPUT ↑', padx=20, pady=5, borderwidth=5, command=copyToClipboardOutput)
button_copy_output_boudy = Button(root, text= f'← Copy OUTPUT', padx=37, pady=10, borderwidth=5, command=copyToClipboardoutput_boudy)
button_dotaz_boudy_v_databazi = Button(root, text= f'↑ Jsou tam? ↑. ', padx=42, pady=0, borderwidth=5, command=dotaz_bouda_v_databazi)
button_cely_obsah_cele_databaze = Button(root, text= f'← Zobrazit celou databázi. ', padx=10, pady=10, borderwidth=5, command=obsah_cele_databaze)

zadatPnLabel.grid(row=0, column=0, columnspan=1, padx=10, pady=10)
entry.grid(row=1, column=0, columnspan=1, padx=10, pady=10)
button_clear.grid(row=2, column=0, padx=5, pady=5)

button_run.grid(row=1, column=1, padx=5, pady=5)
button_boudy.grid(row=2, column=1, padx=5, pady=5)

outputLabel.grid(row=0, column=2, columnspan=1, padx=10, pady=10)
output.grid(row=1, column=2, rowspan=1, columnspan=1, padx=10, pady=10)
button_copy_output.grid(row=2, column=2, padx=5, pady=5)

databaze_check_label.grid(row=3, column=2, padx=10, pady=10)
entry_boudy.grid(row=4, column=2, padx=5, pady=5)
button_dotaz_boudy_v_databazi.grid(row=5, column=2, columnspan=3, padx=10, pady=5)
button_cely_obsah_cele_databaze.grid(row=6, column=2, columnspan=3, padx=10, pady=5)
button_copy_output_boudy.grid(row=7, column=2, padx=10, pady=5)
empty1.grid(row=8, column=2, padx=10, pady=0)

output_boudy.grid(row=3, column=0, rowspan=6, columnspan=2, padx=10, pady=10)





# output_boudy.grid(row=3, column=0, columnspan=2, padx=10, pady=10)





root.mainloop()