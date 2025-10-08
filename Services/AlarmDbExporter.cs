using log4net;
using SIM_TIA_DeviceAlarmgenerator.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SIM_TIA_DeviceAlarmgenerator.Services
{
    /// <summary>
    /// Exportiert eine generierte Alarm-Datenstruktur als Siemens-DB-Quelldatei
    /// im TIA-kompatiblen Textformat (.db).
    /// </summary>
    /// <remarks>
    /// Diese Implementierung erstellt die Verzeichnisstruktur bei Bedarf automatisch
    /// und schreibt die erzeugte Textdatei asynchron in UTF-8 (ohne BOM).
    /// </remarks>
    public sealed class AlarmDbExporter : IAlarmDbExporter
    {
        /// <summary>
        /// Logger-Instanz für Exportinformationen und Fehlermeldungen.
        /// </summary>
        private static readonly ILog Log = App.GetLogger();

        /// <summary>
        /// Exportiert das übergebene Alarmdatenmodell in eine TIA-kompatible .db-Textdatei.
        /// </summary>
        /// <param name="model">Das zu exportierende Datenmodell des Alarm-Datenbausteins.</param>
        /// <param name="filePath">Zielpfad der zu erstellenden .db-Datei.</param>
        /// <param name="ct">Optionaler Abbruch-Token für asynchrone Operationen.</param>
        /// <returns>
        /// Eine <see cref="Task"/>-Instanz, die den asynchronen Schreibvorgang repräsentiert.
        /// </returns>
        /// <remarks>
        /// Erstellt das Zielverzeichnis bei Bedarf automatisch und verwendet UTF-8-Kodierung ohne BOM.
        /// </remarks>
        /// <exception cref="IOException">Wenn beim Schreiben der Datei ein Fehler auftritt.</exception>
        /// <exception cref="UnauthorizedAccessException">Wenn keine Schreibberechtigung für das Zielverzeichnis besteht.</exception>
        /// <exception cref="OperationCanceledException">Wenn der Vorgang über den CancellationToken abgebrochen wird.</exception>
        public async Task ExportAsync(AlarmDbModel model, string filePath, CancellationToken ct = default)
        {
            var text = AlarmDbTextBuilder.Build(model);

            var dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            await File.WriteAllTextAsync(filePath, text, new UTF8Encoding(false), ct).ConfigureAwait(false);

            Log.Info($"alarm.db exportiert → {filePath} (Bytes: {Encoding.UTF8.GetByteCount(text)})");
        }
    }
}
