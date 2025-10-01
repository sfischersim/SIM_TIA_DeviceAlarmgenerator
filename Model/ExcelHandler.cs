using log4net;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SIM_TIA_DeviceAlarmgenerator.Model
{
    internal class ExcelHandler
    {
        public ILog log = App.GetLogger();

        /// <summary>
        /// Liest einen rechteckigen Zellbereich aus einem angegebenen Excel-Worksheet
        /// und gibt alle Zellen als eindimensionales String-Array zurück.
        /// </summary>
        /// <param name="filePath">
        /// Vollständiger Pfad zur Excel-Datei (.xls oder .xlsx).
        /// </param>
        /// <param name="sheetName">
        /// Name des auszulesenden Tabellenblatts.  
        /// Falls der Name nicht existiert, wird eine Ausnahme ausgelöst.
        /// </param>
        /// <param name="startRow">
        /// Erste Zeile des auszulesenden Bereichs (1-basiert).
        /// </param>
        /// <param name="endRow">
        /// Letzte Zeile des auszulesenden Bereichs (1-basiert, inklusiv).
        /// </param>
        /// <param name="startCol">
        /// Erste Spalte des auszulesenden Bereichs (1-basiert, A=1).
        /// </param>
        /// <param name="endCol">
        /// Letzte Spalte des auszulesenden Bereichs (1-basiert, inklusiv).
        /// </param>
        /// <returns>
        /// Ein String-Array mit den gelesenen Zellinhalten.  
        /// Die Werte werden zeilenweise hintereinander gespeichert:  
        /// z. B. erst alle Zellen der ersten Zeile, dann alle der zweiten, usw.  
        /// Leere Zellen werden als leere Strings (<c>""</c>) eingefügt.
        /// </returns>
        /// <remarks>
        /// - Unterstützt sowohl alte Excel-Dateien (.xls) als auch moderne (.xlsx).  
        /// - Verwendet NPOI (<see cref="HSSFWorkbook"/> und <see cref="XSSFWorkbook"/>).  
        /// - Der Rückgabewert ist linearisiert; falls du die Daten lieber
        ///   als zweidimensionales Array (<c>string[row, col]</c>) benötigst,
        ///   müsste die Methode angepasst werden.
        /// </remarks>
        public string[] ReadRangeToArray(
            string filePath,
            string sheetName,
            int startRow,
            int endRow,
            int startCol,
            int endCol)
        {
            var values = new List<string>();

            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook;
                if (Path.GetExtension(filePath).ToLower() == ".xls")
                    workbook = new HSSFWorkbook(fs);  // alte Excel-Version
                else
                    workbook = new XSSFWorkbook(fs); // neue Excel-Version (.xlsx)

                // gewünschtes Tabellenblatt holen
                var sheetIndex = workbook.GetSheetIndex(sheetName);
                if (sheetIndex < 0)
                    throw new ArgumentException($"Sheet '{sheetName}' nicht gefunden in Datei {filePath}.");

                var sheet = workbook.GetSheetAt(sheetIndex);

                for (int r = startRow - 1; r < endRow; r++)   // NPOI ist 0-basiert
                {
                    var row = sheet.GetRow(r);
                    if (row == null) continue;

                    for (int c = startCol - 1; c < endCol; c++)
                    {
                        var cell = row.GetCell(c);
                        values.Add(cell?.ToString() ?? "");
                    }
                }
            }

            return values.ToArray();
        }

        /// <summary>
        /// Macht aus einem flachen string[] eine DataTable mit rows x cols.
        /// </summary>
        public static DataTable FlatToDataTable(string[] flat, int rows, int cols)
        {
            var dt = new DataTable();
            for (int c = 0; c < cols; c++)
                dt.Columns.Add(ColName(c), typeof(string));

            int idx = 0;
            for (int r = 0; r < rows; r++)
            {
                var row = dt.NewRow();
                for (int c = 0; c < cols; c++)
                {
                    row[c] = idx < flat.Length ? flat[idx] : "";
                    idx++;
                }
                dt.Rows.Add(row);
            }
            return dt;
        }

        // Spaltennamen A, B, C, …
        private static string ColName(int index)
        {
            // 0 -> A, 1 -> B, ...
            int n = index;
            string s = "";
            do
            {
                s = (char)('A' + (n % 26)) + s;
                n = n / 26 - 1;
            } while (n >= 0);
            return s;
        }
    }

}
