# FuckYouTransfers
El script del tesorero
Cal tenir python 3 instalat amb la llibreria openpyxl (per editar excels)
Per fer-lo anar omplir un document excel amb els següents headers.
   ___________________________________________________________________________________________
  |_____A____|__B __|___C___|____D___||__E___|___F___|____G____|____H_____|___I____|____J____|
1 |__________Provided Fields_________||___________________Input Fields_______________________|
2 |_Movement_|_Date_|_+Info_|_Import_||_Name_|_Event_|_Concept_|_Advance?_|_Origin_|_Comment_|

A advance cal posar Y o N si forma part d'un avanç i a origin CAIXA o PAYPAL.
Amb Event in Concept anar amb compte de ser coherents i fer servir sempre els mateixos noms.

Per executar cal fer servir:
python3 src/ScriptCuentas.py --file FILEPATH --grants "LIST OF GRANT RELATED EVENTS" --taxes "LIST OF TAX RELATED EVENTS"

La llista de events de grants i tax es refereix al nom del event que s'ha posat com a input field. Per exemple, pel cas present:

python3 src/ScriptCuentas.py --file data/Movimientos_cuenta_0010520.csv --grants "['Grants']" --taxes "['Impuestos']"

Al final demana que s'introduexi la cantitat de diners que hi havia inicialment i al final de mandat al banc 
i al paypal (EN CENTIMS!!)

Un cop acabat l'Actiu i el Passiu s'han de fer manualment però es podria implementar tranquilament donant-li 4 peces de info més.


Accentuo que l'script és una super beta que he fet quan he tingut temps l'estiu de 2019 i és mooooolt millorable en quant a estructura i
manera de fer certes coses. Tot i així, it WORKS!

Un cop jo ja he fet la matada de fer això si es fa servir i es va intentant millorar calculo que la feina del tresorer de contabilitat
es pot veure disminuida dramàticament.

