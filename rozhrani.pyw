from doctest import OutputChecker
from faulthandler import disable
from msilib import type_binary
from posixpath import split
import funkce_prace
import os
import _tkinter
from logging import root
from tkinter import RAISED, Button, Entry, Label, StringVar, Tk, END, Text

# Samotne dotazovani

path_kusovniky_databaze = 'C:\\Users\\Ondrej.rott\\Documents\\Python\\Nove nacenovani\\Programy\\databaze boud s kusovniky.txt'
path_program_databaze = 'C:\\Users\\Ondrej.rott\\Documents\\Python\\Nove nacenovani\\Programy\\seznam programu.txt'

databaze_pro_dotaz = funkce_prace.nacteni_databaze_boud_pro_dotaz(path_kusovniky_databaze)
print("Data nactena a pripravena pro dotazovani...")

kvp_programy = funkce_prace.programy_boud(path_program_databaze)
print("Databaze prgramu vsech boud nacetena...\n")

# Samotne dotazovani

root = Tk()

root.title("UKAZOVADLO PROGRAMU")

zadatPn = StringVar()
zadatPnLabel = Label(root,textvariable=zadatPn, font=('Calibry 10'))
zadatPn.set(f'↓ Tady zadej PN, nebo seznam PN z excelu pod sebou ↓')

entry = Text(root, width = 15, height=5, borderwidth = 5, font=('Calibry 14'))
output = Text(root, width = 25, height=5, borderwidth = 5,font=('Calibry 14'))
output['state'] = 'disabled'
boudy = Text(root,width=60, padx=50 ,borderwidth = 5,height=20, font=('Calibry 10'))
boudy['state'] = 'disabled'


def getProgram():
    output['state'] = 'normal'
    output.delete(1.0, END)
    dotaz = entry.get(1.0, END).split("\n")
    dotaz = [pn.strip() for pn in dotaz if pn != ""]    
    if len(dotaz) != 0:
        vysledek_dotazu = [funkce_prace.dotaz_pn_program(pn, databaze_pro_dotaz, kvp_programy) for pn in dotaz if pn != ""]
        for vysledek in vysledek_dotazu:
            output.insert(END, f'{vysledek[0]}\n')
    output['state'] = 'disabled'

def getBoudy():
    boudy['state'] = 'normal'
    boudy.delete(1.0, END)
    dotaz = entry.get(1.0, END).split("\n")
    dotaz = [pn.strip() for pn in dotaz if pn != ""]    
    if len(dotaz) != 0:
        vysledek_dotazu = [funkce_prace.dotaz_pn_program(pn, databaze_pro_dotaz, kvp_programy) for pn in dotaz if pn != ""]      
        for vysledek in vysledek_dotazu:
            if len(vysledek[1]) != 0: 
                boudy.insert(END, str(vysledek[0].split(":")[0])+" v boudach: "+str(vysledek[1]).replace("'","")+"\n")
            else:
                boudy.insert(END, f'{vysledek[0].split(":")[0]} v boudach: [NELZE URCIT]\n')   
    boudy['state'] = 'disabled'

def clear():
    entry.delete(1.0, END)
    
    output['state'] = 'normal'
    output.delete(1.0, END)
    output['state'] = 'disabled'

    boudy['state'] = 'normal'
    boudy.delete(1.0, END)
    boudy['state'] = 'disabled'

def copyToClipboardOutput():
    x = output.get(1.0, END)
    output.clipboard_clear()
    output.clipboard_append(x)

def copyToClipboardBoudy():
    y = boudy.get(1.0, END)
    boudy.clipboard_clear()
    boudy.clipboard_append(y)

button_run = Button(root, text= f'Urcit program →', padx=20, pady=5, borderwidth=5, command=getProgram)
button_clear = Button(root, text= "CLEAR", padx=20, pady=5, borderwidth=5, command=clear)
button_boudy = Button(root, text= f'Obsazeno v boudach:', padx=20, pady=5, borderwidth=5, command=getBoudy)
button_copy_output = Button(root, text= f'Copy OUTPUT ↑', padx=20, pady=5, borderwidth=5, command=copyToClipboardOutput)
button_copy_boudy = Button(root, text= f'Copy BOUDY\n←', padx=10, pady=10, borderwidth=5, command=copyToClipboardBoudy)

zadatPnLabel.grid(row=0, column=0, columnspan=1, padx=10, pady=10)
entry.grid(row=1, column=0, columnspan=1, padx=10, pady=10)
output.grid(row=1, column=2, columnspan=1, padx=10, pady=10)
button_run.grid(row=1, column=1, padx=5, pady=5)
button_clear.grid(row=2, column=0, padx=5, pady=5)
button_boudy.grid(row=2, column=1, padx=5, pady=5)
button_copy_output.grid(row=2, column=2, padx=5, pady=5)
boudy.grid(row=3, column=0, columnspan=2, padx=10, pady=10)
button_copy_boudy.grid(row=3, column=2, columnspan=3, padx=10, pady=10)

root.mainloop()