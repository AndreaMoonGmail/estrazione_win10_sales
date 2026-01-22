package com.example.estrazione_win10_sales.service;

import com.example.estrazione_win10_sales.model.CsvRecord;

import java.time.Instant;
import java.util.*;
import java.util.stream.Collectors;

public class TerminalColumnService {

    private static final class ValTs {
        final String value;
        final Instant ts;
        ValTs(String value, Instant ts) {
            this.value = value;
            this.ts = ts;
        }
    }

    /**
     * Compute terminals per riceId. For each riceId, produce a list of terminal values (terminal_1..terminal_N)
     * following rules from DEVELOPMENT_PLAN.md.
     * This version selects for each tid the most recent osType/osVersion based on CsvRecord.timestamp.
     */
    public Map<String, List<String>> computeTerminalsByRiceId(List<CsvRecord> records) {
        if (records == null) return Collections.emptyMap();

        // group by riceId
        Map<String, List<CsvRecord>> byRice = records.stream()
                .filter(r -> r.getRiceId() != null && !r.getRiceId().isBlank())
                .collect(Collectors.groupingBy(CsvRecord::getRiceId, LinkedHashMap::new, Collectors.toList()));

        Map<String, List<String>> result = new LinkedHashMap<>();

        for (Map.Entry<String, List<CsvRecord>> e : byRice.entrySet()) {
            String riceId = e.getKey();
            List<CsvRecord> recs = e.getValue();

            // group records by tid and collect latest osType/osVersion for that tid (by timestamp)
            Map<String, Map<String, ValTs>> tidMeta = new HashMap<>();
            for (CsvRecord r : recs) {
                String tid = r.getTid();
                if (tid == null || tid.isBlank()) continue;
                Map<String, ValTs> meta = tidMeta.computeIfAbsent(tid, k -> new HashMap<>());
                String key = r.getChiave();
                if (key == null) continue;
                String val = r.getValore();
                if (val == null) continue;
                Instant ts = r.getTimestamp();
                ValTs existing = meta.get(key);
                if (existing == null) {
                    meta.put(key, new ValTs(val, ts));
                } else {
                    Instant existingTs = existing.ts;
                    // decide replacement: prefer non-null newer timestamp; if both null keep existing (first seen)
                    if (ts != null) {
                        if (existingTs == null || ts.isAfter(existingTs)) {
                            meta.put(key, new ValTs(val, ts));
                        }
                    }
                    // if ts == null and existingTs == null -> keep existing; if existingTs != null and ts==null -> keep existing
                }
            }

            // determine ordering of tids: try numeric sort, fallback to lexical
            List<String> tids = new ArrayList<>(tidMeta.keySet());
            tids.sort((a, b) -> {
                try {
                    Long la = Long.parseLong(a);
                    Long lb = Long.parseLong(b);
                    return la.compareTo(lb);
                } catch (NumberFormatException ex) {
                    return a.compareTo(b);
                }
            });

            List<String> terminals = new ArrayList<>();
            for (String tid : tids) {
                Map<String, ValTs> meta = tidMeta.get(tid);
                String osType = null;
                String osVersion = null;
                if (meta != null) {
                    ValTs osTypeT = meta.get("osType");
                    ValTs osVersionT = meta.get("osVersion");
                    if (osTypeT != null && osTypeT.value != null) osType = osTypeT.value;
                    if (osVersionT != null && osVersionT.value != null) osVersion = osVersionT.value;
                }
                String val;
                if ((osType == null || osType.isBlank()) && (osVersion == null || osVersion.isBlank())) {
                    // tid present but no metadata -> mark as N.D.
                    val = "N.D.";
                } else if (osType != null && osType.equalsIgnoreCase("WINDOWS")) {
                    if (osVersion != null && osVersion.startsWith("10.0.2")) {
                        val = tid + "_WIN11";
                    } else {
                        val = tid + "_WIN10";
                    }
                } else {
                    // for non-WINDOWS we return null so the Excel writer will leave the cell empty
                    val = null;
                }
                terminals.add(val);
            }

            // Diagnostic: if this is the riceId under investigation, print tids and terminals
            if ("283840".equals(riceId)) {
                System.out.println("DEBUG-TerminalColumnService: riceId=283840 tids=" + tids + " terminals=" + terminals);
            }

            result.put(riceId, terminals);
        }

        return result;
    }

    public int maxTerminalCount(Map<String, List<String>> terminalsByRiceId) {
        if (terminalsByRiceId == null || terminalsByRiceId.isEmpty()) return 0;
        return terminalsByRiceId.values().stream().mapToInt(List::size).max().orElse(0);
    }
}
