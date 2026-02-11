@echo off
echo ===== COMPILAZIONE E TEST APPLICAZIONE =====
echo.

cd "C:\Users\sartirana\Documents\sorgenti_git\estrazione_win10_sales"

echo [1] Compilazione in corso...
call mvn clean package -DskipTests

if %errorlevel% neq 0 (
    echo ERRORE: Compilazione fallita!
    pause
    exit /b 1
)

echo.
echo [2] Esecuzione applicazione...
echo Output file atteso: target\estrazione_output.xlsx
echo.

call java -jar target\estrazione_win10_sales-0.0.1-SNAPSHOT.jar

echo.
echo [3] Verifica risultati...
if exist "target\estrazione_output.xlsx" (
    echo ✅ SUCCESS: File estrazione_output.xlsx generato correttamente!
    echo Controlla che abbia 2360 righe nel file target\estrazione_output.xlsx
) else (
    echo ❌ ERROR: File estrazione_output.xlsx non trovato
)

echo.
echo ===== TEST COMPLETATO =====
pause
