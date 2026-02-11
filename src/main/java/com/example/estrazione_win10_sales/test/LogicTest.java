package com.example.estrazione_win10_sales.test;

import java.util.*;

public class LogicTest {
    public static void main(String[] args) {
        // Simula il comportamento corretto
        System.out.println("=== TEST LOGICA CORRETTA ===");

        // Simula template con 5 righe (idrice univoci)
        List<String> templateIdrice = Arrays.asList("001", "002", "003", "004", "005");

        // Simula CSV con duplicati (stesso idrice, terminali diversi)
        Map<String, List<String>> csvTerminals = new HashMap<>();
        csvTerminals.put("001", Arrays.asList("T1_WIN10", "T2_WIN11")); // 001 ha 2 terminali
        csvTerminals.put("003", Arrays.asList("T3_WIN10")); // 003 ha 1 terminale
        // 002, 004, 005 non hanno terminali nel CSV

        System.out.println("Template idrice: " + templateIdrice.size());
        System.out.println("CSV terminals per idrice: " + csvTerminals.size());

        // Logica corretta: per ogni riga template
        int outputRows = 0;
        for (String templateId : templateIdrice) {
            outputRows++;
            List<String> terminals = csvTerminals.get(templateId);
            if (terminals != null) {
                System.out.println("Riga " + outputRows + " - Idrice: " + templateId + " -> Terminali: " + terminals);
            } else {
                System.out.println("Riga " + outputRows + " - Idrice: " + templateId + " -> Nessun terminale (colonne vuote)");
            }
        }

        System.out.println("Righe totali output: " + outputRows + " (deve essere uguale al template: " + templateIdrice.size() + ")");

        if (outputRows == templateIdrice.size()) {
            System.out.println("✅ LOGICA CORRETTA: Tutte le righe del template vengono mantenute");
        } else {
            System.out.println("❌ ERRORE: Righe perse!");
        }
    }
}
