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
        /// und baut daraus ein vollständiges <see cref="AlarmDbModel"/>.
        /// </summary>
        /// <param name="filePath">Vollständiger Pfad zur Excel-Datei.</param>
        /// <returns>
        /// Ein initialisiertes <see cref="AlarmDbModel"/> mit Standard- und Applikationsalarmen,
        /// gefüllter <c>DeviceSetting</c>-Matrix, <c>DeviceTypes</c>, berechnetem <c>DevicesWordCount</c>
        /// und optionalen Instanzlabels aus dem Sheet <c>Devices</c>.
        /// </returns>
        /// <exception cref="InvalidOperationException">
        /// Das erwartete Sheet <c>StA-Cfg</c> wurde nicht gefunden.
        /// </exception>
        /// <exception cref="NotSupportedException">
        /// Die Dateierweiterung wird nicht unterstützt.
        /// </exception>
        /// <remarks>
        /// - Excel wird im Sharing-Modus geöffnet (<see cref="FileShare.ReadWrite"/>).<br/>
        /// - Bei ungültigen/fehlenden Zahlenwerten werden 0 bzw. sinnvolle Minimalwerte verwendet.<br/>
        /// - <c>DevicesWordCount</c> basiert auf <c>sum(ceil(AlmQty/16) * DevQty)</c>, mindestens 1.
        /// </remarks>
        public AlarmDbModel BuildFromExcel(string filePath)
        {
            // Öffne die Excel-Datei im Lesemodus.
            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            // Erzeuge ein NPOI-Workbook-Objekt abhängig vom Dateityp.
            IWorkbook wb = CreateWorkbook(fs, Path.GetExtension(filePath));

            // Erzeuge ein neues AlarmDbModel mit allen Standard-Alarmbits.
            var model = CreateBaseModelWithStd(wb);

            // Lade das Tabellenblatt „StA-Cfg“ (Gerätetypenliste).
            var sta = wb.GetSheet(SHEET_STA)
                ?? throw new InvalidOperationException($"Sheet '{SHEET_STA}' nicht gefunden.");

            // Daten beginnen hinter der Headerzeile.
            int dataStart = STA_HEADER_ROW + 1;

            // Anzahl der vorhandenen Gerätetypen ermitteln
            int deviceTypes = 0;
            for (int r = dataStart; r <= sta.LastRowNum; r++)
            {
                var row = sta.GetRow(r);
                if (row == null) continue;

                var devName = GetString(row.GetCell(COL_DEVNAME)).Trim();
                if (!string.IsNullOrWhiteSpace(devName))
                    deviceTypes++;
            }

            // DeviceSetting konfigurieren
            model.DeviceSettingRows = Math.Max(1, deviceTypes);
            model.DeviceSettingCols = 3;
            model.DeviceSetting = new uint[model.DeviceSettingRows, model.DeviceSettingCols];
            model.DeviceSettingMax = model.DeviceSettingRows;

            // DeviceSetting füllen & Words zählen
            long wordsTotal = 0;
            int logicalIdx = 0;

            for (int r = dataStart; r <= sta.LastRowNum; r++)
            {
                var row = sta.GetRow(r);
                if (row == null) continue;

                var devName = GetString(row.GetCell(COL_DEVNAME)).Trim();

                // Wenn Zelle mit Device-Bezeichnung leer ist -> mit nächster Zeile fortfahren
                if (string.IsNullOrWhiteSpace(devName)) 
                    continue;

                logicalIdx++;

                var qty = ToUInt(row.GetCell(COL_DEVQTY));
                var alm = ToUInt(row.GetCell(COL_ALMQTY));

                model.DeviceTypes.Add(new DeviceTypeInfo
                {
                    DevName = devName,
                    DevQty = (int)qty,
                    AlmQty = (int)alm
                });

                // Matrix: [0] = Geräteanzahl, [1] = Alarmanzahl
                model.DeviceSetting[logicalIdx - 1, 0] = qty;
                model.DeviceSetting[logicalIdx - 1, 1] = alm;

                if (qty > 0 && alm > 0)
                {
                    var wordsPerDevice = (int)Math.Ceiling(alm / 16.0);
                    wordsTotal += (long)qty * wordsPerDevice;
                }
            }

            // Ergebnisse ins Model schreiben
            model.DevicesWordCount = (int)Math.Max(1, wordsTotal);

            // Optional: zusätzliche Informationen aus dem Sheet „Devices“
            TryReadDevicesSheet(wb, model);

            return model;
        }

        /// <summary>
        /// Erstellt ein NPOI-Workbook passend zur Dateierweiterung.
        /// </summary>
        /// <param name="s">Geöffneter Stream der Excel-Datei.</param>
        /// <param name="ext">Dateierweiterung inkl. Punkt (z. B. „.xlsx“).</param>
        /// <returns>Ein initialisiertes <see cref="IWorkbook"/>.</returns>
        /// <exception cref="NotSupportedException">Wenn <paramref name="ext"/> nicht unterstützt wird.</exception>
        private static IWorkbook CreateWorkbook(Stream s, string ext)
            => ext.ToLowerInvariant() switch
            {
                ".xls" => new HSSFWorkbook(s),
                ".xlsx" or ".xlsm" => new XSSFWorkbook(s),
                _ => throw new NotSupportedException($"Erweiterung nicht unterstützt: {ext}")
            };

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
