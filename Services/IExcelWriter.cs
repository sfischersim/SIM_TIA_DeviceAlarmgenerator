using System.Threading;
using System.Threading.Tasks;

namespace SIM_TIA_DeviceAlarmgenerator.Services
{
    /// <summary>
    /// Schreibt Tabellen (2D-Matrizen) als Excel-Dateien.
    /// </summary>
    /// <remarks>
    /// Implementierungen sollten eine UTF-8-kompatible .xlsx-Arbeitsmappe erzeugen und den Zielordner bei Bedarf anlegen.
    /// </remarks>
    public interface IExcelWriter
    {
        /// <summary>
        /// Speichert eine String-Matrix als Excel-Arbeitsmappe (<c>.xlsx</c>).
        /// </summary>
        /// <param name="data">Matrix <c>[Zeile, Spalte]</c> mit Zellinhalten. Leere Strings sind erlaubt.</param>
        /// <param name="usedRows">
        /// Anzahl der zu schreibenden Zeilen aus <paramref name="data"/>. 
        /// Werte &lt;= 0 erzeugen eine leere Datei (nur Header, falls implementierungsseitig vorhanden).
        /// Wird auf die tatsächliche Datenlänge begrenzt.
        /// </param>
        /// <param name="sheetName">Name des Arbeitsblatts. Empfohlen: „Alarms“.</param>
        /// <param name="outputFolder">Zielordner; wird bei Bedarf erstellt.</param>
        /// <param name="fileName">Dateiname (z. B. „Alarms.xlsx“).</param>
        /// <param name="autoSizeColumns">Ob die Spaltenbreite automatisch an den Inhalt angepasst werden soll.</param>
        /// <param name="ct">Abbruchtoken zur vorzeitigen Beendigung.</param>
        /// <returns>Vollständiger Dateipfad der erzeugten Datei.</returns>
        /// <remarks>
        /// Die Implementierung sollte große Matrizen streaming-fähig schreiben und den Abbruchtoken regelmäßig prüfen.
        /// </remarks>
        /// <exception cref="System.ArgumentNullException">Wenn <paramref name="data"/> oder <paramref name="sheetName"/> oder <paramref name="outputFolder"/> <see langword="null"/> ist.</exception>
        /// <exception cref="System.ArgumentException">Wenn <paramref name="fileName"/> leer ist oder ungültige Pfadzeichen enthält.</exception>
        /// <exception cref="System.IO.IOException">Fehler beim Schreiben der Datei.</exception>
        /// <exception cref="System.UnauthorizedAccessException">Keine Schreibberechtigung im Zielordner.</exception>
        /// <exception cref="System.OperationCanceledException">Der Vorgang wurde über <paramref name="ct"/> abgebrochen.</exception>
        Task<string> SaveMatrixAsync(
            string[,] data,
            int usedRows,
            string sheetName,
            string outputFolder,
            string fileName = "Alarms.xlsx",
            bool autoSizeColumns = true,
            CancellationToken ct = default);
    }
}
