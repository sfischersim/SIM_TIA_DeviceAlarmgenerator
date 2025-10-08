using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace SIM_TIA_DeviceAlarmgenerator.Services
{
    /// <summary>
    /// NPOI-basierte Implementierung zum Speichern von 2D-Matrizen als <c>.xlsx</c>.
    /// </summary>
    public sealed class NpoiExcelWriter : IExcelWriter
    {
        /// <inheritdoc />
        public async Task<string> SaveMatrixAsync(
            string[,] data,
            int usedRows,
            string sheetName,
            string outputFolder,
            string fileName = "Alarms.xlsx",
            bool autoSizeColumns = true,
            CancellationToken ct = default)
        {
            if (data is null) throw new ArgumentNullException(nameof(data));
            if (string.IsNullOrWhiteSpace(sheetName)) sheetName = "Alarms";
            if (string.IsNullOrWhiteSpace(outputFolder)) outputFolder = Environment.CurrentDirectory;

            int rows = Math.Min(Math.Max(0, usedRows), data.GetLength(0));
            int cols = data.GetLength(1);

            Directory.CreateDirectory(outputFolder);
            string path = Path.Combine(outputFolder, fileName);

            using var wb = new XSSFWorkbook();

            // Basis-Schrift und Stil
            var font = (XSSFFont)wb.CreateFont();
            font.FontHeightInPoints = 11;
            font.FontName = "Calibri";

            var style = (XSSFCellStyle)wb.CreateCellStyle();
            style.SetFont(font);

            // Blatt anlegen
            var sheet = wb.CreateSheet(sheetName);

            // Zeilen/Zellen schreiben
            for (int r = 0; r < rows; r++)
            {
                ct.ThrowIfCancellationRequested();
                var row = sheet.CreateRow(r);

                for (int c = 0; c < cols; c++)
                {
                    var cell = row.CreateCell(c, CellType.String);
                    cell.SetCellValue(data[r, c] ?? string.Empty);
                    cell.CellStyle = style;
                }
            }

            // Optional: Spaltenbreite automatisch
            if (autoSizeColumns && rows > 0)
            {
                for (int c = 0; c < cols; c++)
                {
                    try { sheet.AutoSizeColumn(c); }
                    catch { /* manche Inhalte lassen AutoSize scheitern – ignorieren */ }
                }
            }

            // Persistieren
            await using var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
            wb.Write(fs);

            return path;
        }
    }
}
