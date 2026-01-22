package com.example.estrazione_win10_sales.model;

import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Objects;

public class CsvRecord {
    private final String riceId;
    private final String chiave;
    private final String tid;
    private final String valore;
    private final Instant timestamp; // nullable if not parsable or absent

    public CsvRecord(String riceId, String chiave, String tid, String valore, Instant timestamp) {
        this.riceId = riceId;
        this.chiave = chiave;
        this.tid = tid;
        this.valore = valore;
        this.timestamp = timestamp;
    }

    public String getRiceId() {
        return riceId;
    }

    public String getChiave() {
        return chiave;
    }

    public String getTid() {
        return tid;
    }

    public String getValore() {
        return valore;
    }

    public Instant getTimestamp() {
        return timestamp;
    }

    /**
     * Construct a CsvRecord from a CSV row (String[]).
     * Supports both formats:
     * - [riceId, chiave, tid, valore, timestamp]
     * - [riceId, chiave, tid, timestamp_TO_DELETE, valore, timestamp]
     */
    public static CsvRecord fromCsvRow(String[] row) {
        if (row == null || row.length < 4) {
            throw new IllegalArgumentException("CSV row must have at least 4 columns: riceId, chiave, tid, valore");
        }
        String r = sanitize(row[0]);
        String k = null;
        String t = null;
        String v = null;
        String tsRaw = null;

        // Detect layout where second column duplicates riceId (header shows: rice_id;rice_id;chiave;terminal_id;...)
        if (row.length >= 7 && sanitize(row[1]) != null && sanitize(row[1]).equals(sanitize(row[0]))) {
            // layout: 0=rice_id,1=rice_id,2=chiave,3=terminal_id,4=timestamp_TO_DELETE,5=valore,6=timestamp
            k = sanitize(row[2]);
            t = sanitize(row[3]);
            v = sanitize(row[5]);
            tsRaw = sanitize(row[6]);
        } else {
            // Fallback heuristics for other layouts
            k = sanitize(row[1]);
            t = sanitize(row[2]);
            if (row.length >= 7) {
                v = sanitize(row[5]);
                tsRaw = sanitize(row[6]);
            } else if (row.length == 6) {
                v = sanitize(row[4]);
                tsRaw = sanitize(row[5]);
            } else if (row.length == 5) {
                v = sanitize(row[3]);
                tsRaw = sanitize(row[4]);
            } else if (row.length == 4) {
                v = sanitize(row[3]);
            }
        }

        Instant ts = null;
        if (tsRaw != null && !tsRaw.isBlank()) {
            // try parsing common format: "yyyy-MM-dd HH:mm:ss" or ISO-like
            String cleaned = tsRaw.trim();
            // remove surrounding quotes if any
            if (cleaned.startsWith("\"") && cleaned.endsWith("\"")) {
                cleaned = cleaned.substring(1, cleaned.length() - 1);
            }
            // Try ISO first
            try {
                ts = Instant.parse(cleaned);
            } catch (DateTimeParseException ex1) {
                // Try 'yyyy-MM-dd HH:mm:ss'
                try {
                    DateTimeFormatter f = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
                    LocalDateTime ldt = LocalDateTime.parse(cleaned, f);
                    ts = ldt.atZone(ZoneId.systemDefault()).toInstant();
                } catch (DateTimeParseException ex2) {
                    // give up, leave null
                }
            }
        }
        return new CsvRecord(r, k, t, v, ts);
    }

    private static String sanitize(String s) {
        if (s == null) return null;
        String t = s.trim();
        if (t.startsWith("\"") && t.endsWith("\"") && t.length() >= 2) {
            t = t.substring(1, t.length() - 1);
        }
        if (t.isBlank()) return null;
        if (t.equalsIgnoreCase("NULL")) return null;
        return t;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof CsvRecord)) return false;
        CsvRecord that = (CsvRecord) o;
        return Objects.equals(riceId, that.riceId) && Objects.equals(chiave, that.chiave) && Objects.equals(tid, that.tid) && Objects.equals(valore, that.valore) && Objects.equals(timestamp, that.timestamp);
    }

    @Override
    public int hashCode() {
        return Objects.hash(riceId, chiave, tid, valore, timestamp);
    }
}
