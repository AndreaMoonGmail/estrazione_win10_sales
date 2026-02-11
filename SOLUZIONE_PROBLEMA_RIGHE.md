# SOLUZIONE AL PROBLEMA: Perdita di righe dal template.xlsx

## PROBLEMA IDENTIFICATO
- **Template.xlsx**: 2360 righe (incluso header)
- **Output attuale**: 1923 righe 
- **Righe perse**: 437 righe

## CAUSA DEL PROBLEMA
Il codice originale aveva una logica SBAGLIATA:
1. Iterava sui riceId trovati nel CSV
2. Per ogni riceId del CSV, cercava la riga corrispondente nel template
3. Le righe del template senza corrispondenza nel CSV venivano IGNORATE

## SOLUZIONE IMPLEMENTATA
Ho modificato la logica in `CsvToExcelService.java`:

### PRIMA (SBAGLIATO):
```java
// Per ogni riceId nel CSV...
for (Map.Entry<String, List<String>> e : terminalsByRiceId.entrySet()) {
    String riceId = e.getKey(); // riceId dal CSV
    // Cerca la riga nel template
    // Se non trova, SALTA (perdita di riga)
}
```

### DOPO (CORRETTO):
```java
// Per ogni riga nel TEMPLATE...
for (int r = 1; r <= lastRowNum; r++) {
    Row templateRow = sheet.getRow(r);
    String templateRiceId = ExcelUtils.getStringCellValue(templateRow.getCell(0));
    
    // Cerca i terminali per questo riceId nel CSV
    List<String> terminals = terminalsByRiceId.get(templateRiceId.trim());
    
    if (terminals != null && !terminals.isEmpty()) {
        // Aggiungi le colonne terminale
        for (int i = 0; i < terminals.size(); i++) {
            String header = "terminal_" + (i + 1);
            // Popola la cella con il valore del terminale
        }
    } else {
        // Nessun terminale - lascia le colonne vuote (CORRETTO!)
    }
}
```

## COMPORTAMENTO CORRETTO
1. **Mantiene tutte le 2360 righe del template**
2. **Per ogni riga del template**:
   - Se trova terminali nel CSV → aggiunge colonne `terminal_1`, `terminal_2`, etc.
   - Se NON trova terminali → lascia le colonne vuote (come richiesto)
3. **Gestisce correttamente i duplicati nel CSV**: se un idrice ha più terminali nel CSV, vengono tutti aggiunti

## ESEMPIO DI DUPLICATO NEL CSV
Dal log precedente, l'idrice **285551** è un esempio che puoi controllare manualmente:
- Cerca "285551" nel file `template.xlsx` 
- Cerca "285551" nel file `estrazione.csv`
- Vedrai che nel CSV può comparire più volte (terminali diversi)

## FILE MODIFICATI
- `src/main/java/com/example/estrazione_win10_sales/service/CsvToExcelService.java`
- Rimossa la logica di rimozione duplicati dal template
- Rimossa la logica di mapping CSV → template
- Aggiunta nuova logica template → CSV

## RISULTATO ATTESO
- **File output**: 2360 righe (stesso numero del template)
- **Colonne terminale**: popolate solo per gli idrice che hanno corrispondenza nel CSV
- **Righe senza terminali**: mantengono colonne terminale vuote

La soluzione è stata implementata e dovrebbe risolvere completamente il problema.
