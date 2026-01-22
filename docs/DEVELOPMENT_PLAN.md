## Phase 1
- read column A (Idrice) of inputfile/template.xlsx
- search in estrazione.csv for matching rice_id values
- when value matches

1 make a copy of template.xlsx  named result.xlsx and put them into same folder of template.xlsx
in this new file do following operations:
- for each rice_id found in estrazione.csv do the following operations:
- check osType column value

1 if osType is equals to WINDOWS
- add a new colm for each different tid whith same rice_id
- the new colums name will be: terminal_x where x starts with 1 and increments by 1. For exaple terminal_1, terminal_2, terminal_3 ecc...
- if osVersion starts with 10.0.2 for eah new terminal colums the value will be: tid value concat _ WIN11

- 2 if osType is not equals to 'WINDOWS'
- print in the new terminal colum N.D.
- for example I can hav terminal_1 = N.D., terminal_2 = WIN10, terminal_3 = WIN11
- the primted value depemds on the osType value and on the osVersion value thi the rules described above
if there are two tid with same timestamp use recently one.

3 rimuovi tutte le righe con riceid multiple,lasciando i valori delle colonne terminal_1, terminal_2 ecc... sulla riga con osVersion, eliminando quindi la riga con osType.
Quindi mi aspetto per ogni riceid al piu una riga, che ha da 0 a n terminal_x colonne a seconda di quante righe con lo stesso riceid sono state trovate in estrazione.csv