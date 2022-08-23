***Funkce:***

- Účelem programu Prevody_aplikace je projít všechny prodejní linky z PZN105 ve stanoveném rozpětí dnů a zkontrolovat, jestli bude v den
	prodeje dostatečné množství požadované položky na skladu PZN105.
- Pokud ano, vráti informaci, že linka je v pořádku. 
- Pokud ne, prověří pro stejnou položku také sklad PZN100 a pokud je možno převést požadované mmožství z PZN100, vráti informaci o možnosti převodu (ANO/NE).

***Logika programu:***

1. Program si načte prodejní linky z CQ reportu z LN "webreports\After Sales\Ondra test\TEST_Sales_order_lines_master_plan.eq".
	• prověřují se linky, které:
		a) jsou "Shortage". Podmínka je zde, že inventory on hand dané linky minus suma všech 
			ordered qty stejného itemu na linkách předcházejících této lince je menší než 0.
		b) obsahují pouze neprojektové položky.
		c) mají alespoň jeden z datumů Planned delivery date, Customer requested date a Supplier promised date v časovém rozmezí, které se prověřuje (zadané uživatelem).

2. Program si načte data Item order plánů PZN100 a PZN105 z CQ reportu "webreports\After Sales\Ondra test\PZN100+105_order_plan.eq" pro itemy z prověřovaných linek z kroku 1 .
	• z načtených dat si sestaví order plány PZN100 a PZN105 daných položek.

3. Program prochází jednotlivé prodejní linky z kroku 1 a k danému datumu Planned delivery date prověřuje Planned available množství itemu z linky
   na skladech PZN100 a PZN105 ze sestavených order plánů.
   	• Pokud je v den prodeje Planned delivery date dostatečné Planned available množství itemu na PZN105:
		→ linka je vyhodnocena jako v pořádku.
	• Pokud není v den prodeje Planned delivery date dostatečné Planned available množství itemu na PZN105:
		→ program se dívá, jestli je v den prodeje Planned delivery date dostatečné Planned available množství itemu na PZN100:
			• Pokud je: 
				→ podívá se, jestli by se převodem na PZN105 neohrozily ostatní linky PZN100. (Kontrola je, ze pripadne nakup stihne znovu objednat - pouziva purchase LT z Purchase dat + 5 pracovnich dnu jako safety time).
					• Pokud by se převodem neohrozily ostatní linky PZN100:
						→ linka se vyhodnotí jako převod z PZN100 na PZN105.
			• Pokud není:
				→ linka se vyhodnotí jako že nelze udělat převod z PZN100 na PZN105.
	
	POZN.: Při prověřování order plánů se ignorují linky Planned Purchase orders. (berou se jako by tam nebyly)