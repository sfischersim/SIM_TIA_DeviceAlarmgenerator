using System;

namespace SIM_TIA_DeviceAlarmgenerator.Model
{
    /// <summary>
    /// Definiert Methoden zum Einlesen von Excel-Datenquellen (z. B. „StA-Cfg“, „Devices“, „AppAlarm“)
    /// in strukturierte 2D-Arrays für die weitere Verarbeitung.
    /// </summary>
    public interface IExcelSource
    {
        /// <summary>
        /// Liest ein angegebenes Tabellenblatt als zweidimensionales String-Array.
        /// </summary>
        /// <param name="filePath">Pfad zur Excel-Datei (.xlsx oder .xlsm).</param>
        /// <param name="sheetName">Name des auszulesenden Tabellenblatts.</param>
        /// <param name="maxCols">
        /// Optionale maximale Spaltenanzahl.  
        /// Wenn angegeben, werden nur diese Spalten berücksichtigt.
        /// </param>
        /// <returns>
        /// Ein zweidimensionales String-Array <c>[Zeile, Spalte]</c>, das die Zellinhalte enthält.
        /// Leere Zellen werden als <see cref="string.Empty"/> zurückgegeben.
        /// </returns>
        /// <exception cref="ArgumentException">Wenn das angegebene Tabellenblatt nicht existiert.</exception>
        string[,] ReadSheet(string filePath, string sheetName, int? maxCols = null);

        /// <summary>
        /// Liest das Blatt <c>Devices</c> aus einer Excel-Datei und gibt dessen Inhalte als Matrix zurück.
        /// </summary>
        /// <param name="filePath">Pfad zur Excel-Datei.</param>
        /// <param name="maxCols">Maximale Spaltenanzahl (Standard: 35).</param>
        /// <returns>Ein zweidimensionales String-Array mit den Inhalten des „Devices“-Sheets.</returns>
        string[,] ReadDevices(string filePath, int maxCols = 35);

        /// <summary>
        /// Liest das Blatt <c>StA-Cfg</c> aus und gibt die ermittelten Geräte- und Alarmmengen zurück.
        /// </summary>
        /// <param name="filePath">Pfad zur Excel-Datei.</param>
        /// <returns>
        /// Ein zweidimensionales Integer-Array <c>[30, 2]</c>:
        /// <list type="bullet">
        /// <item><description>Spalte 0: Geräteanzahl (Excel-Spalte C)</description></item>
        /// <item><description>Spalte 1: Alarmanzahl pro Gerät (Excel-Spalte D)</description></item>
        /// </list>
        /// </returns>
        /// <exception cref="ArgumentException">Wenn das Tabellenblatt nicht existiert.</exception>
        int[,] ReadNumberOfDevices(string filePath);

        /// <summary>
        /// Liest das Blatt <c>AppAlarm (_AA)</c> aus einer Excel-Datei als String-Matrix.
        /// </summary>
        /// <param name="filePath">Pfad zur Excel-Datei.</param>
        /// <param name="maxCols">Maximale Spaltenanzahl (Standard: 22).</param>
        /// <returns>Ein zweidimensionales String-Array mit den Inhalten des AppAlarm-Blatts.</returns>
        string[,] ReadAppAlarm(string filePath, int maxCols = 22);

        /// <summary>
        /// Liest das Blatt <c>StA-Cfg</c> als String-Matrix (inklusive Typnamen in Spalte 1).  
        /// Diese Methode ist speziell für die Mengenermittlung pro Device-Typ vorgesehen.
        /// </summary>
        /// <param name="filePath">Pfad zur Excel-Datei (.xlsx oder .xlsm).</param>
        /// <returns>
        /// Eine 2D-Matrix <c>[Zeilen, Spalten]</c> mit allen Zellinhalten als Text.
        /// </returns>
        string[,] ReadStaCfgText(string filePath);
    }
}
