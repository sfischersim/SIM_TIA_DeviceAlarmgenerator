using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SIM_TIA_DeviceAlarmgenerator.Model;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace SIM_TIA_DeviceAlarmgenerator.Services
{
    /// <summary>
    /// Liest StA-Cfg (und optional „Devices“, „ApplicationAlarms“) aus einer Excel-Datei
    /// und baut daraus ein vollständiges <see cref="AlarmDbModel"/> für den Export.
    /// </summary>
    /// <remarks>
    /// Unterstützte Formate: <c>.xls</c>, <c>.xlsx</c>, <c>.xlsm</c> (NPOI).
    /// Aktuell werden DeviceSetting, DevicesWordCount sowie die Typ- und Instanzinformationen abgedeckt.
    /// </remarks>
    public sealed class ExcelAlarmConfigReader
    {
        /// <summary>Tabellenblattname mit der Standardkonfiguration.</summary>
        private const string SHEET_STA = "StA-Cfg";

        /// <summary>Tabellenblattname mit den Geräteinstanzen.</summary>
        private const string SHEET_DEVICES = "Devices";

        /// <summary>Tabellenblattname mit den Geräteinstanzen.</summary>
        private const string SHEET_APPALARM = "AppAlarm (_AA)";

        /// <summary>0-basierter Index der englischen Headerzeile im Sheet <c>StA-Cfg</c>.</summary>
        private const int STA_HEADER_ROW = 1; // englische Headerzeile

        /// <summary>Spaltenindex „Idx“ im Sheet <c>StA-Cfg</c> (0-basiert).</summary>
        private const int COL_IDX = 0;        // Idx

        /// <summary>Spaltenindex „DevName“ im Sheet <c>StA-Cfg</c> (0-basiert).</summary>
        private const int COL_DEVNAME = 1;    // DevName

        /// <summary>Spaltenindex „DevQty“ im Sheet <c>StA-Cfg</c> (0-basiert).</summary>
        private const int COL_DEVQTY = 2;    // DevQty

        /// <summary>Spaltenindex „AlmQty“ im Sheet <c>StA-Cfg</c> (0-basiert).</summary>
        private const int COL_ALMQTY = 3;    // AlmQty

        // Spalten & Header definieren (0-basiert)
        private const int APP_HEADER_ROWS = 2;  // 2 Kopfzeilen
        private const int APP_COL_NAME = 2;  // Name/Key
        private const int APP_COL_CLASS = 3;  // "3" => Warnings, sonst Errors
        private const int APP_COL_TEXT = 7;  // Kommentar/Meldungstext

        /// <summary>
        /// Liest die Alarmkonfigurationsdaten aus einer Excel-Datei ein
        /// und baut daraus ein vollständiges <see cref="AlarmDbModel"/>-Objekt.
        /// </summary>
        /// <param name="filePath">
        /// Vollständiger Pfad zur Excel-Datei (.xlsx, .xls oder .xlsm).
        /// </param>
        /// <returns>
        /// Ein initialisiertes <see cref="AlarmDbModel"/>, das alle
        /// erforderlichen Header- und Gerätekonfigurationsinformationen enthält.
        /// Zusätzlich werden die Inhalte der relevanten Sheets als
        /// 2-dimensionale Text-Matrizen (<c>DevicesMatrix</c>,
        /// <c>StaCfgTextMatrix</c> und <c>AppAlarmMatrix</c>) für die UI-Preview bereitgestellt.
        /// </returns>
        /// <exception cref="InvalidOperationException">
        /// Wird ausgelöst, wenn das erwartete Sheet <c>StA-Cfg</c> in der Excel-Datei
        /// nicht gefunden wurde.
        /// </exception>
        /// <exception cref="NotSupportedException">
        /// Wird ausgelöst, wenn die Dateierweiterung nicht unterstützt wird.
        /// </exception>
        /// <remarks>
        /// <list type="bullet">
        /// <item>
        /// Die Excel-Datei wird im <see cref="FileShare.ReadWrite"/>-Modus geöffnet,
        /// sodass sie parallel in Excel geöffnet bleiben kann.
        /// </item>
        /// <item>
        /// Es werden die Sheets <c>Devices</c>, <c>StA-Cfg</c> und
        /// <c>AppAlarm (_AA)</c> (bzw. deren Varianten) gesucht. 
        /// Die Suche erfolgt tolerant gegenüber abweichenden Schreibweisen
        /// (z. B. Leer-, Unter- oder Bindestriche).
        /// </item>
        /// <item>
        /// Enthält ein Sheet ungültige oder leere Zellen, werden diese
        /// als leere Zeichenfolgen (<c>string.Empty</c>) in die Matrix übernommen.
        /// </item>
        /// <item>
        /// Die erzeugten Textmatrizen dienen ausschließlich der Anzeige
        /// und Vorschau in der UI (<c>DevicesPreview</c>, <c>NumbersPreview</c>,
        /// <c>AppAlarmPreview</c>) und werden im <see cref="AlarmDbModel"/> abgelegt.
        /// </item>
        /// <item>
        /// Der eigentliche fachliche Aufbau des Alarm-Modells (Standard- und
        /// Applikationsalarme, DeviceTypes, DeviceInstanceLabels usw.)
        /// erfolgt nach dem Einlesen dieser Rohdaten.
        /// </item>
        /// </list>
        /// </remarks>
        public AlarmDbModel BuildFromExcel(string filePath)
        {
            // --- 1. Workbook öffnen ---
            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            IWorkbook wb = Path.GetExtension(filePath).Equals(".xls", StringComparison.OrdinalIgnoreCase)
                ? new HSSFWorkbook(fs)
                : new XSSFWorkbook(fs); // liest auch .xlsm

            // --- 2. Hilfsfunktionen lokal definieren ---
            static ISheet? FindSheet(IWorkbook wb, params string[] candidates)
            {
                foreach (var name in candidates)
                {
                    var sheet = wb.GetSheet(name);
                    if (sheet != null) return sheet;
                }
                // fallback: tolerant
                var wanted = candidates.Select(c => c.ToLowerInvariant().Replace(" ", "").Replace("-", "").Replace("_", ""));
                for (int i = 0; i < wb.NumberOfSheets; i++)
                {
                    var s = wb.GetSheetAt(i);
                    var norm = s.SheetName.ToLowerInvariant().Replace(" ", "").Replace("-", "").Replace("_", "");
                    if (wanted.Any(w => norm.Contains(w))) return s;
                }
                return null;
            }

            static string[,] ReadSheetToMatrix(ISheet? sheet)
            {
                if (sheet == null)
                    return new string[0, 0];

                int firstRow = sheet.FirstRowNum;
                int lastRow = sheet.LastRowNum;

                int cols = 0;
                for (int r = firstRow; r <= lastRow; r++)
                    cols = Math.Max(cols, sheet.GetRow(r)?.LastCellNum ?? 0);
                if (cols <= 0)
                    return new string[0, 0];

                int rows = (lastRow - firstRow) + 1;
                var matrix = new string[rows, cols];
                var fmt = new DataFormatter(System.Globalization.CultureInfo.GetCultureInfo("de-DE"));

                for (int r = firstRow; r <= lastRow; r++)
                {
                    var row = sheet.GetRow(r);
                    for (int c = 0; c < cols; c++)
                    {
                        var cell = row?.GetCell(c);
                        matrix[r - firstRow, c] = cell == null ? string.Empty : fmt.FormatCellValue(cell);
                    }
                }
                return matrix;
            }

            // --- 3. Sheets finden + Matrizen lesen ---
            // Ggf. weitere mögliche Sheet-Bezeichnungn anhängen, z.B. var sheetAppAlm = FindSheet(wb, "AppAlarm (_AA)", "AppAlarm", "AppAlarmAA");

            var sheetDevices = FindSheet(wb, "Devices");
            var sheetStaCfg = FindSheet(wb, "StA-Cfg");
            var sheetAppAlm = FindSheet(wb, "AppAlarm (_AA)");

            string[,] devicesMatrix = ReadSheetToMatrix(sheetDevices);
            string[,] staCfgTextMatrix = ReadSheetToMatrix(sheetStaCfg);
            string[,] appAlarmMatrix = ReadSheetToMatrix(sheetAppAlm);

            // --- 4. Bestehende Logik: Model aufbauen ---
            var model = new AlarmDbModel();

            model = CreateBaseModelWithStd(wb);

            TryReadStaCfgSheet(wb, model);
            TryReadDevicesSheet(wb, model);

            model.DevicesMatrix = devicesMatrix;
            model.StaCfgTextMatrix = staCfgTextMatrix;
            model.AppAlarmMatrix = appAlarmMatrix;

            return model;
        }



        /// <summary>
        /// Liest eine Zelle als String (robust gegenüber unterschiedlichen Zelltypen).
        /// </summary>
        /// <param name="c">Zellreferenz.</param>
        /// <returns>Getrimmter Text oder ein leerer String.</returns>
        private static string GetString(ICell? c)
            => c == null ? "" :
               c.CellType switch
               {
                   CellType.String => c.StringCellValue ?? "",
                   CellType.Numeric => c.NumericCellValue.ToString(CultureInfo.InvariantCulture),
                   CellType.Boolean => c.BooleanCellValue ? "TRUE" : "FALSE",
                   CellType.Formula => c.ToString(),
                   _ => ""
               };

        /// <summary>
        /// Konvertiert eine Zelle in einen <see cref="uint"/>-Wert (negativ → 0, leer → 0).
        /// </summary>
        /// <param name="c">Zellreferenz.</param>
        /// <returns>Konvertierter Wert oder 0 bei Ungültigkeit.</returns>
        private static uint ToUInt(ICell? c)
        {
            if (c == null) return 0;
            if (c.CellType == CellType.Numeric && c.NumericCellValue >= 0) return (uint)c.NumericCellValue;
            if (c.CellType == CellType.String &&
                uint.TryParse(c.StringCellValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out var u)) return u;
            return 0;
        }

        /// <summary>
        /// Erstellt ein neues <see cref="AlarmDbModel"/> mit Standard-Alarmbits
        /// und ergänzt optional Applikationsalarme aus dem Sheet „AppAlarm (_AA)“.
        /// </summary>
        /// <param name="wb">Optional: geöffnetes Workbook zum Einlesen von „AppAlarm (_AA)“.</param>
        /// <returns>Ein vorkonfiguriertes <see cref="AlarmDbModel"/>.</returns>
        /// <remarks>
        /// Standardalarme bilden den festen Grundstock des DB-Layouts; Applikationsalarme sind optional.
        /// </remarks>
        private static AlarmDbModel CreateBaseModelWithStd(IWorkbook? wb = null)
        {
            var m = new AlarmDbModel();

            // Standard-Alarmbits (Systemalarme)
            m.StandardAlarms.AddRange(new[]
            {
                new AlarmBit { Name="HMI_AL_DevTypOOR", AlarmText="collectAlarm: Fehlender oder falscher DevTyp beim Aufruf verwendet" },
                new AlarmBit { Name="HMI_AL_QtyBit", AlarmText="collectAlarm: Anzahl Fehlerbits sind außerhalb parametrierten Bereich, DevTyp deaktiviert" },
                new AlarmBit { Name="HMI_AL_NoInit", AlarmText="collectAlarm: HMI_Alarme Init muss durchgeführt werden" },
                new AlarmBit { Name="HMI_AL_DevTypCfg", AlarmText="collectAlarm: Fehler in Device Konfiguration. Devicetyp #" },
                new AlarmBit { Name="HMI_AL_DevIdx", AlarmText="collectAlarm: Device Index ist außerhalb parametrierten Bereichs" },
                new AlarmBit { Name="HMI_AL_devArrMax", AlarmText="collectAlarm: aktueller Device Index außerhalb gültigen Bereichs" },
                new AlarmBit { Name="HMI_AL_Fu", AlarmText="collectAlarm: falscher Funktionsaufruf" },
                new AlarmBit { Name="PCS_Err", AlarmText="PCS: Allgemeiner Fehler in Kommunikation" },
                new AlarmBit { Name="PCS_Err_TA", AlarmText="PCS: Telegrammfehler Nr #, Transaktionsüberwachung" },
                new AlarmBit { Name="PCS_TO_Cycle", AlarmText="PCS: Neue Daten bevor letzte Übertragung beendet. Laufzeit-Fehler Kommunikation" },
                new AlarmBit { Name="PCS_ErrSetData", AlarmText="PCS: Fehler in SetData" },
                new AlarmBit { Name="PCS_ErrGetData", AlarmText="PCS: Fehler in GetData" },
                new AlarmBit { Name="P_ArrStartIdx", AlarmText="P oder P_Adv Start-Index falsch" },
                new AlarmBit { Name="Res15", AlarmText="Reserviert (Platzhalter)" },
                new AlarmBit { Name="PCS_ErrRcvRcp", AlarmText="PCS: Fehler beim Laden einer Rezeptur" },
                new AlarmBit { Name="NoJobData", AlarmText="PCS: keine gültigen Auftragsdaten" },
                new AlarmBit { Name="NoRcpData", AlarmText="PCS: keine gültigen Rezepturdaten" },
                new AlarmBit { Name="PD_Ctrl_Error", AlarmText="PD: Sammelfehler (Teiledatenmanager)" }
            });

            // Optionale Applikationsalarme aus „AppAlarm (_AA)“
            if (wb != null)
            {
                var sh = wb.GetSheet(SHEET_APPALARM);
                if (sh != null)
                {
                    for (int r = APP_HEADER_ROWS; r <= sh.LastRowNum; r++)
                    {
                        var row = sh.GetRow(r);
                        if (row == null) continue;

                        string name = GetString(row.GetCell(APP_COL_NAME)).Trim();
                        string comment = GetString(row.GetCell(APP_COL_TEXT)).Trim();
                        string clsCode = GetString(row.GetCell(APP_COL_CLASS)).Trim();

                        // ganz leere Zeilen überspringen
                        if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(comment)) continue;

                        // Klassenmapping wie im Alt-Tool
                        string @class = (clsCode == "3") ? "Warnings" : "Errors";

                        m.ApplicationAlarms.Add(new AlarmBit
                        {
                            Name = name,
                            AlarmText = comment,
                            Class = @class
                        });
                    }
                }
            }

            return m;
        }

        /// <summary>
        /// Reads the StA-Cfg sheet and fills the alarm database model with device settings.
        /// </summary>
        /// <param name="wb">The opened Excel workbook.</param>
        /// <param name="model">The alarm database model that receives the device settings.</param>
        /// <remarks>
        /// The Excel column "Idx" is the single source of truth for the generated DeviceSetting index.
        /// Missing Idx values are kept as empty DeviceSetting rows with zero values.
        /// </remarks>
        private void TryReadStaCfgSheet(IWorkbook wb, AlarmDbModel model)
        {
            var sh = wb.GetSheet(SHEET_STA);
            if (sh == null)
                return;

            int dataStart = STA_HEADER_ROW + 1;
            int maxDeviceTypeIdx = 0;

            #region Determine DeviceSetting Size

            for (int r = dataStart; r <= sh.LastRowNum; r++)
            {
                var row = sh.GetRow(r);
                if (row == null)
                    continue;

                string devName = GetString(row.GetCell(COL_DEVNAME)).Trim();
                if (string.IsNullOrWhiteSpace(devName))
                    continue;

                uint idx = ToUInt(row.GetCell(COL_IDX));
                if (idx == 0)
                    continue;

                maxDeviceTypeIdx = Math.Max(maxDeviceTypeIdx, (int)idx);
            }

            #endregion

            #region Configure DeviceSetting Matrix

            model.DeviceSettingRows = Math.Max(1, maxDeviceTypeIdx);
            model.DeviceSettingCols = 3;
            model.DeviceSetting = new uint[model.DeviceSettingRows, model.DeviceSettingCols];
            model.DeviceSettingMax = model.DeviceSettingRows;

            #endregion

            #region Fill DeviceSetting Matrix

            long wordsTotal = 0;

            for (int r = dataStart; r <= sh.LastRowNum; r++)
            {
                var row = sh.GetRow(r);
                if (row == null)
                    continue;

                string devName = GetString(row.GetCell(COL_DEVNAME)).Trim();
                if (string.IsNullOrWhiteSpace(devName))
                    continue;

                uint idx = ToUInt(row.GetCell(COL_IDX));
                if (idx == 0)
                    continue;

                uint qty = ToUInt(row.GetCell(COL_DEVQTY));
                uint alm = ToUInt(row.GetCell(COL_ALMQTY));

                int deviceSettingRowIndex = (int)idx - 1;

                model.DeviceTypes.Add(new DeviceTypeInfo
                {
                    DevName = devName,
                    DevQty = (int)qty,
                    AlmQty = (int)alm
                });

                // DeviceSetting uses the Excel Idx as the authoritative position.
                // Missing Idx values remain initialized with 0, 0, 0.
                model.DeviceSetting[deviceSettingRowIndex, 0] = qty;
                model.DeviceSetting[deviceSettingRowIndex, 1] = alm;
                model.DeviceSetting[deviceSettingRowIndex, 2] = 0;

                if (qty > 0 && alm > 0)
                {
                    int wordsPerDevice = (int)Math.Ceiling(alm / 16.0);
                    wordsTotal += (long)qty * wordsPerDevice;
                }
            }

            #endregion

            #region Write Results To Model

            model.DevicesWordCount = (int)Math.Max(1, wordsTotal);

            #endregion
        }
        /// <summary>
        /// Liest aus dem Sheet „Devices“ BMK-Gruppen und Instanznamen pro Gerätetyp
        /// und füllt <see cref="AlarmDbModel.DeviceInstanceLabels"/>.
        /// </summary>
        /// <param name="wb">Geöffnetes Workbook.</param>
        /// <param name="model">Zu befüllendes <see cref="AlarmDbModel"/>.</param>
        /// <remarks>
        /// Erkennt blockweise Struktur: Headerzeile (Typ/BMK/Name), Aliaszeile, anschließend Datenzeilen.
        /// Ein Block endet bei leerer oder ungültiger Zeile.
        /// </remarks>
        private void TryReadDevicesSheet(IWorkbook wb, AlarmDbModel model)
        {
            var sh = wb.GetSheet(SHEET_DEVICES);
            if (sh == null) return;

            int r = sh.FirstRowNum;

            while (r + 2 <= sh.LastRowNum)
            {
                var hdr = sh.GetRow(r);
                if (hdr == null)
                {
                    r++;
                    continue;
                }

                string c1 = (hdr.GetCell(1)?.ToString() ?? "").Trim().ToLowerInvariant();
                string c2 = (hdr.GetCell(2)?.ToString() ?? "").Trim().ToLowerInvariant();
                string c3 = (hdr.GetCell(3)?.ToString() ?? "").Trim().ToLowerInvariant();

                bool looksHeader =
                    c1 == "bmk gruppe" &&
                    c2 == "bmk symbol" &&
                    c3.Contains("name des devices sprache 1");

                if (!looksHeader)
                {
                    r++;
                    continue;
                }

                string devType = (hdr.GetCell(0)?.ToString() ?? "").Trim();
                if (string.IsNullOrEmpty(devType))
                {
                    r++;
                    continue;
                }

                int rr = r + 2;
                var instMap = new Dictionary<int, DeviceInstanceLabel>();

                for (; rr <= sh.LastRowNum; rr++)
                {
                    var row = sh.GetRow(rr);
                    if (row == null)
                        break;

                    bool isEmpty =
                        string.IsNullOrWhiteSpace(row.GetCell(0)?.ToString()) &&
                        string.IsNullOrWhiteSpace(row.GetCell(1)?.ToString()) &&
                        string.IsNullOrWhiteSpace(row.GetCell(3)?.ToString());

                    if (isEmpty)
                        break;

                    if (!int.TryParse(row.GetCell(0)?.ToString(), out int instIdx))
                        break;

                    string bmk = (row.GetCell(1)?.ToString() ?? "").Trim();
                    string nameS1 = (row.GetCell(3)?.ToString() ?? "").Trim();

                    instMap[instIdx] = new DeviceInstanceLabel
                    {
                        BmkGroup = bmk,
                        NameS1 = nameS1
                    };
                }

                if (instMap.Count > 0)
                    model.DeviceInstanceLabels[devType] = instMap;

                r = rr + 1;
            }
        }
    }
}
