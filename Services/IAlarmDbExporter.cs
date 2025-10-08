using SIM_TIA_DeviceAlarmgenerator.Model;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace SIM_TIA_DeviceAlarmgenerator.Services
{
    /// <summary>
    /// Schnittstelle für den Export von alarm.db-Dateien (TIA-kompatibles DB-Quelltextformat).
    /// </summary>
    /// <remarks>
    /// Implementierungen sollten Verzeichnisse bei Bedarf anlegen und in UTF-8 (vorzugsweise ohne BOM) schreiben.
    /// </remarks>
    public interface IAlarmDbExporter
    {
        /// <summary>
        /// Erzeugt aus dem Modell eine .db-Quelltextdatei und schreibt sie auf die Platte.
        /// </summary>
        /// <param name="model">Das Datenmodell des Alarm-Datenbausteins.</param>
        /// <param name="filePath">Zieldateipfad der zu erstellenden .db-Datei.</param>
        /// <param name="ct">Optionaler Abbruch-Token für asynchrone Operationen.</param>
        /// <returns>Eine <see cref="Task"/>-Instanz, die den asynchronen Exportvorgang repräsentiert.</returns>
        /// <exception cref="ArgumentNullException">Wenn <paramref name="model"/> oder <paramref name="filePath"/> <c>null</c> ist.</exception>
        /// <exception cref="IOException">Fehler beim Schreiben der Datei.</exception>
        /// <exception cref="UnauthorizedAccessException">Keine Schreibberechtigung im Zielverzeichnis.</exception>
        /// <exception cref="OperationCanceledException">Der Vorgang wurde über den <paramref name="ct"/> abgebrochen.</exception>
        Task ExportAsync(AlarmDbModel model, string filePath, CancellationToken ct = default);
    }
}
