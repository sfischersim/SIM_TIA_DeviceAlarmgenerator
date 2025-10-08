using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Globalization;
using System.IO;

namespace SIM_TIA_DeviceAlarmgenerator.Model
{
    #region Implementation

    /// <summary>
    /// Implementiert das Einlesen von Excel-Dateien (XLSX/XLSM) mit NPOI.
    /// Unterstützt die Blätter <c>Devices</c>, <c>StA-Cfg</c>, <c>AppAlarm (_AA)</c>
    /// sowie generische Blätter.
    /// </summary>
    [Obsolete]
    public sealed class NpoiExcelSource : IExcelSource
    {
        /// <summary>
        /// Liest ein angegebenes Tabellenblatt als zweidimensionales String-Array.
        /// </summary>
        /// <param name="filePath">Pfad zur Excel-Datei (.xlsx/.xlsm).</param>
        /// <param name="sheetName">Name des auszulesenden Blatts.</param>
        /// <param name="maxCols">
        /// Optionale maximale Spaltenanzahl. Wenn angegeben, werden nur diese Spalten berücksichtigt.
        /// </param>
        /// <returns>
        /// Ein zweidimensionales String-Array <c>[Zeile, Spalte]</c>, das die Inhalte des Blatts enthält.
        /// Leere Zellen werden als <see cref="string.Empty"/> zurückgegeben.
        /// </returns>
        /// <exception cref="ArgumentException">Wenn das Blatt nicht existiert.</exception>
        public string[,] ReadSheet(string filePath, string sheetName, int? maxCols = null)
        {
            using var fs = File.OpenRead(filePath);
            using var wb = new XSSFWorkbook(fs);

            var idx = wb.GetSheetIndex(sheetName);
            if (idx < 0)
                throw new ArgumentException($"Sheet '{sheetName}' nicht gefunden in '{filePath}'.");

            var sheet = wb.GetSheetAt(idx);

            // NPOI: LastRowNum ist Index der letzten belegten Zeile (0-basiert)
            var rowCount = sheet.LastRowNum + 1;
            if (rowCount <= 0) return new string[0, 0];

            // Maximale Spalten ermitteln
            int computedMaxCols = 0;
            for (int r = 0; r <= sheet.LastRowNum; r++)
            {
                var row = sheet.GetRow(r);
                if (row == null) continue;
                // LastCellNum ist eine Count (1-basiert), daher direkt als Breite nutzbar
                computedMaxCols = Math.Max(computedMaxCols, row.LastCellNum);
            }

            int cols = computedMaxCols;
            if (maxCols.HasValue)
                cols = Math.Min(cols, maxCols.Value);

            var data = new string[rowCount, cols];

            for (int r = 0; r < rowCount; r++)
            {
                var row = sheet.GetRow(r);
                if (row == null) continue;

                var lastCellNum = Math.Min(row.LastCellNum, (short)cols);
                for (int c = 0; c < lastCellNum; c++)
                {
                    var cell = row.GetCell(c);
                    data[r, c] = cell is null ? string.Empty : GetCellString(cell);
                }
                // Rest bleibt string.Empty (Default des Arrays ist null → aber wir füllen nur bis lastCellNum)
            }

            return data;
        }

        /// <summary>
        /// Liest das Blatt <c>Devices</c> aus einer Excel-Datei als Matrix.
        /// </summary>
        /// <param name="filePath">Pfad zur Excel-Datei.</param>
        /// <param name="maxCols">Maximale Spaltenanzahl (Standard: 35).</param>
        /// <returns>Ein zweidimensionales String-Array für das Devices-Blatt.</returns>
        public string[,] ReadDevices(string filePath, int maxCols = 35)
            => ReadSheet(filePath, "Devices", maxCols);

        /// <summary>
        /// Liest das Blatt <c>StA-Cfg</c> und gibt die Geräte-/Alarmanzahlen zurück.
        /// </summary>
        /// <param name="filePath">Pfad zur Excel-Datei.</param>
        /// <returns>
        /// Ein zweidimensionales Array <c>[30,2]</c>:
        /// <list type="bullet">
        /// <item><description>Spalte 0: Anzahl Geräte (Excel-Spalte C)</description></item>
        /// <item><description>Spalte 1: Anzahl Alarme pro Gerät (Excel-Spalte D)</description></item>
        /// </list>
        /// </returns>
        /// <exception cref="ArgumentException">Wenn das Blatt nicht existiert.</exception>
        public int[,] ReadNumberOfDevices(string filePath)
        {
            using var fs = File.OpenRead(filePath);
            using var wb = new XSSFWorkbook(fs);

            var idx = wb.GetSheetIndex("StA-Cfg");
            if (idx < 0)
                throw new ArgumentException($"Sheet 'StA-Cfg' nicht gefunden in '{filePath}'.");

            var sheet = wb.GetSheetAt(idx);

            // Zielstruktur wie im Altprojekt: [30,2]
            var result = new int[30, 2];

            // Altcode: ab Zeile 2 (0-basiert), also Excel-Zeile 3; Spalten C(2) & D(3)
            for (int r = 2; r <= sheet.LastRowNum && r < 31; r++)
            {
                var row = sheet.GetRow(r);
                if (row == null) continue;

                result[r - 1, 0] = TryParseInt(row.GetCell(2));
                result[r - 1, 1] = TryParseInt(row.GetCell(3));
            }

            return result;
        }

        /// <summary>
        /// Liest das Blatt <c>AppAlarm (_AA)</c> aus einer Excel-Datei als Matrix.
        /// </summary>
        /// <param name="filePath">Pfad zur Excel-Datei.</param>
        /// <param name="maxCols">Maximale Spaltenanzahl (Standard: 22).</param>
        /// <returns>Ein zweidimensionales String-Array für das AppAlarm-Blatt.</returns>
        public string[,] ReadAppAlarm(string filePath, int maxCols = 22)
            => ReadSheet(filePath, "AppAlarm (_AA)", maxCols);

        /// <summary>
        /// Liest das Blatt <c>StA-Cfg</c> als Textmatrix (<c>string[,]</c>) aus einer XLSX/XLSM-Datei.
        /// Nutzt NPOI (<see cref="XSSFWorkbook"/>) und formatiert Zahlen/Datumswerte robust.
        /// </summary>
        /// <param name="filePath">Pfad zur Excel-Datei (.xlsx/.xlsm).</param>
        /// <returns>Matrix [Zeile, Spalte] mit Zellinhalten als String.</returns>
        /// <exception cref="ArgumentException">Wenn <paramref name="filePath"/> leer ist.</exception>
        /// <exception cref="InvalidOperationException">Wenn das Sheet <c>StA-Cfg</c> nicht existiert.</exception>
        public string[,] ReadStaCfgText(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("Pfad zur Excel-Datei fehlt.", nameof(filePath));

            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var wb = new XSSFWorkbook(fs); // kann xlsm

            var idx = wb.GetSheetIndex("StA-Cfg");
            if (idx < 0)
                throw new InvalidOperationException("Sheet \"StA-Cfg\" wurde nicht gefunden.");

            var sheet = wb.GetSheetAt(idx);
            var lastRow = sheet.LastRowNum;
            int rows = Math.Max(1, lastRow + 1);

            // Spaltenbreite heuristisch bestimmen (oberes Limit genügt, z. B. 16)
            int maxCols = 16;
            int detectedCols = 0;
            for (int r = 0; r <= lastRow; r++)
            {
                var row = sheet.GetRow(r);
                if (row is null) continue;
                detectedCols = Math.Max(detectedCols, row.LastCellNum);
                if (detectedCols >= maxCols) { detectedCols = maxCols; break; }
            }
            int cols = Math.Max(1, Math.Min(detectedCols == 0 ? maxCols : detectedCols, maxCols));

            var matrix = new string[rows, cols];

            for (int r = 0; r < rows; r++)
            {
                var row = sheet.GetRow(r);
                if (row is null) continue;

                int limit = Math.Min(cols, row.LastCellNum);
                for (int c = 0; c < limit; c++)
                {
                    matrix[r, c] = GetCellString(row.GetCell(c));
                }
            }

            return matrix;
        }

        /// <summary>
        /// Konvertiert eine NPOI-Zelle in einen nicht-null String.
        /// </summary>
        /// <param name="cell">Die auszulesende Zelle.</param>
        /// <returns>Niemals <c>null</c>; leere Zellen → <see cref="string.Empty"/>.</returns>
        private static string GetCellString(ICell cell)
        {
            if (cell is null) return string.Empty;

            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue ?? string.Empty;

                case CellType.Boolean:
                    return cell.BooleanCellValue ? "TRUE" : "FALSE";

                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        // Tolerant: macht aus dem NPOI-Wert sicher einen DateTime
                        var dt = Convert.ToDateTime(cell.DateCellValue);
                        return dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                    }
                    return cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);

                case CellType.Formula:
                    // NPOI kann null liefern → absichern
                    return cell.ToString() ?? string.Empty;

                case CellType.Blank:
                    return string.Empty;

                default:
                    return cell.ToString() ?? string.Empty;
            }
        }

        /// <summary>
        /// Versucht eine Zelle als Integer zu parsen; gibt bei leeren/ungültigen Zellen <c>0</c> zurück.
        /// </summary>
        /// <param name="cell">Die Zelle (kann null sein).</param>
        /// <returns>Geparster Integer oder <c>0</c>.</returns>
        private static int TryParseInt(ICell? cell)
        {
            if (cell == null) return 0;

            if (cell.CellType == CellType.Numeric && !DateUtil.IsCellDateFormatted(cell))
                return (int)Math.Round(cell.NumericCellValue);

            var s = GetCellString(cell);
            return int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : 0;
        }
    }

    #endregion
}
