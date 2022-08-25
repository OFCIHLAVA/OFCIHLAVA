***Funkce:***

- Účelem programu SFE x BFE je vzít input od uživatele v podobě jednoho nebo více P/N a všechny tyto P/N prověřit, v jakých se nacházejí boudách a jaký je
	program těchto boud. Výsledek vrátí v podobě PN:PROGRAM pro každou položku.


***Příprava***
- Program používá 2 primitivní databáze v podobě textových souborů. Obě by měly zrdcadlově obsahovat data ke stejnému seznamu boud. Pokud by byla nějaká data obsažena pouze v jedné z nich, nevrátilo by to pro ně výsledek.
	
	1. je "seznam programu.txt":
		• Jedná se o .txt soubor, s se seznamem boud a jejich programem (BFE/SFE/MIX). Data byla získaná z PSR excel databáze na jaře 2022 - za jejich správnost se autor nemůže zaručit.
			Data MUSÍ být uložena v následujícím formátu:				
				a) Všechna data jsou pouze na 1 řádku txt souboru.
				b) Jedná se o souvislý řetězec znaků, ve formátu: P/Nboudy1:PROGRAMboudy1,P/Nboudy2:PROGRAMboudy2,P/Nboudy2:PROGRAMboudy2, atd.
					kde:
						• znak :	rozděluje jednotlivé boudy a jejich program.
						• znak ,	značí další boudu
					pozn.: 	Na pořadí boud v databázi nezáleží.
						každá bouda v databázi by měla mít přidělen právě jeden PROGRAM. Pokud by nastala situace, kdy je reálně bouda BFE i SFE, sem do databáze je potřeba ji uložit pouze jednou, jako MIX.
						V datech nesmějí být žádné mezery.

		• "seznam programu.txt" AKTUALIZACE:
			Seznam boud v databázi "seznam programu.txt" by měl odpovídat seznamu všech boud z druhé databáze "databaze boud s kusovniky.txt".
			Pro hromadnou kontrolu je potřeba získat všechna P/N následující před dělícím znakem ":" - všechny unikatni boudy v databázi. (Způsob praktického provedení nechávám zcela na čtenářově fantazii...;) )

			Pro přidání nových boud je třeba přidat data na konec stávajícíh dat ve formátu viz. 1 b).
			Pro smazání starých / nesprávných dat stačí z dat smazat požadované boudy s jejich programem.

	2. je "databaze boud s kusovniky.txt":
		• Jedná se o .txt soubor, s kusovníky vybraných boud z LN. Data jsou získaná z CQ reportu "10lvl_ebom_programy.eq" ve slozce "users/Ondrej Rott/nacenovani/Sales group programy".
			Data MUSÍ být uložena v následujícím formátu:
				a) Všechna data jsou pouze na 1 řádku txt souboru.
				b) Jedná se o souvislý řetězec znaků, ve formátu: <+++>P/Nboudy1:PN/boudy1|Item1boudy1|Item2boudy1|Item3boudy1|....posledniItemboudy1|<+++>P/Nboudy2:PN/boudy2|Item1boudy2|Item2boudy2|Item3boudy2|....posledniItemboudy2| atd.
					kde:
						• znak <+++> 	značí začátek kusovníku nové boudy.
						• znak :	rozděluje jednotlivé itemy kusovníku boudy.
					pozn.: 	Na pořadí dílů v jednotlivých kusovnících boudy nezáleží (kontroluje se pouze přítomnost / nepřítomnost v boudě).
						V jednotlivých kusovnícíh boudy je potřeba odstranit duplicitní itemy (duplicity by něměly vadit funkčnosti programu samotného, ale když by se ponechaly v databázi, ta pak obrovsky nabobtná a
							běh programu bude mnohonásobně zpomalený).
						V datech nesmějí být žádné mezery.
		
		• "databaze boud s kusovniky.txt" AKTUALIZACE:
			Obsah databáze byl vyjetý z LN na jaře 2022. Vyjížděly se boudy, které v té době byly v PSR excelové databázi + nějaké ručně s Petrem Skalou.
			Seznam boud v databázi "databaze boud s kusovniky.txt" by měl odpovídat seznamu všech boud z druhé databáze "seznam programu.txt".
			
			• Ruční aktualizace databáze":
				Pro hromadnou kontrolu je potřeba získat všechna P/N následující hned po dělícím znaku "<+++>" - všechny unikatni boudy v databázi. (Způsob praktického provedení nechávám zcela na čtenářově fantazii...;) ) 
				Pro přidání nových boud je třeba vyjet CQ report pro nové boudy a přidat je nakonec dat ve formátu viz. 2 b) výše.
				Pro smazání starých / nesprávných dat stačí z dat smazat požadované boudy s jejich kusovníkem.
			• Aktualizace databáze pomocí programu "databaze_update_aplikace.py".
				Nejprve je potřeba vyjet CQ report pro kusvoníky boud, které chceme přidat do databáze viz. bod 2. a uložit jako txt soubor do složky "Y:\Departments\Sales and Marketing\Aftersales\11_PLANNING\23_Python_utilities\SFExBFExMIX\databaze\Exporty LN". 	
				Dále je třeba spustit program "databaze_update_aplikace.py" ve složce "Y:\Departments\Sales and Marketing\Aftersales\11_PLANNING\23_Python_utilities\SFExBFExMIX".
				Program by měl porovnat stávajícíc databázi s boudami, které chcceme přidat.
					→ Pokud už v databázi jsou a jejich kusovníky jsou totožné → nic se neupdatuje.
					→ Pokud tam nejsou, nebo jsou, ale jejich kusovníky se neshodují → přidají se / přepíšou ve stávající databázi.
					→ Výsledek se uloží jako soubor "POST_UPDATE_programy_databaze_update.txt" do složky "Y:\Departments\Sales and Marketing\Aftersales\11_PLANNING\23_Python_utilities\SFExBFExMIX\databaze".
					→ Pokud chceme používat dále tuto updatovanou databázi, je třeba přepsat soubor puvodní databáze ve stejné složce tímto novým souborem a změnit jméno na původní - "databaze boud s kusovniky.txt"						

***Logika programu:***
1. Program si vezme postupně každou položku ze seznamu zadaného uživatelem a prověří databázi kusovníků, zda je v ní někde obsažen.
	• Pokud ano"
		a) Vezme seznam P/N všech boud, ve kterých se daná položka vyskytuje a tento seznam prověří s databází boud a programů.
		b) Zde prověří, jaké programy mají dané boudy z vytvořeného seznamu a jejich kombinací určí výsledný program:
			• Pokud je některá z boud MIX → položka je MIX.
			• Jinak pokud je některá položka SFE a žádná jiná není BFE → položka je SFE.
			• Jinak pokud je některá položka BFE a žádná jiná není SFE → položka je BFE.

	Výsledek vrátí jako seznam párů položka:PROGRAM pro všechny dotazované položky.

2. (Volitelné) Pokud je třeba také vidět v jakých boudách(s jejich programem) se dané položky nacházejí. Program vrátí pro každou položku seznam boud, v jejichž kusovníku se položka nachází.
	Pozn.: Toto se děje jako vedlější produkt kroku 1a) výše.

		
