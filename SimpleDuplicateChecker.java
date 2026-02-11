import java.io.*;
import java.util.*;

public class SimpleDuplicateChecker {
    public static void main(String[] args) {
        System.out.println("Analisi semplice del file template.xlsx...");

        // Proviamo un approccio più semplice - contare le righe nel file CSV di estrazione
        try {
            // Leggiamo prima il CSV per capire quanti record abbiamo
            Map<String, Integer> csvCounts = new HashMap<String, Integer>();

            try (BufferedReader br = new BufferedReader(new FileReader("src/main/resources/inputfile/estrazione.csv"))) {
                String line;
                int csvRows = 0;
                while ((line = br.readLine()) != null) {
                    csvRows++;
                    if (csvRows == 1) continue; // skip header

                    String[] parts = line.split(";");
                    if (parts.length > 0) {
                        String riceId = parts[0].trim().replaceAll("\"", "");
                        if (!riceId.isEmpty()) {
                            csvCounts.put(riceId, csvCounts.getOrDefault(riceId, 0) + 1);
                        }
                    }
                }
                System.out.println("Righe totali nel CSV estrazione: " + (csvRows - 1)); // -1 per header
                System.out.println("RiceId univoci nel CSV: " + csvCounts.size());
            }

            System.out.println("\nPer identificare i duplicati nel template.xlsx:");
            System.out.println("1. Apri il file src/main/resources/inputfile/template.xlsx");
            System.out.println("2. Seleziona la colonna A (Idrice)");
            System.out.println("3. Usa 'Formattazione condizionale' -> 'Evidenzia celle duplicate'");
            System.out.println("4. Oppure usa un filtro avanzato per trovare valori duplicati");

            System.out.println("\nIl template dovrebbe avere 2360 righe (incluso header)");
            System.out.println("Se l'output finale ha 1923 righe, significa che vengono rimosse " + (2360 - 1923) + " righe");

        } catch (IOException e) {
            System.err.println("Errore lettura file: " + e.getMessage());
        }
    }
}
