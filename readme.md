
<H1>Doplňky pro MS Excel přes API</H1>

Projekt obsahuje dvě verze 

 - VBA - První verze doplňků napsaná ve Visual Basic for Applications v MS Excel
 - VB.NET - Druhá verze napsaná ve VB .NET - Verze .NET 4.8. Obashuje celý projekt ve Visual Studiu

VB.NET
Pro rozběhnutí projektu ve Visual Studiu je třeba stáhnout si visual studio 2022 comunity edition. Otevřít projekt AbraExcelAddIn.sln a pak provézt build. V adresáři AbraExcelAddIn\AbraExcelAddIn\bin bude buď v debug nebo v release (záleží na tom co jste zvolili pro build) vše potřebné. 
Většinou se používá 32 bitová verze doplňku. Projekt využívá https://excel-dna.net aktuálně ve verzi 1.1.1 a některé části vev verzi 1.1.0 (ty co nemají verzi 1.1.1) - lze zjistit v Nuget repozitáři jaké verze jsou dostupné.
