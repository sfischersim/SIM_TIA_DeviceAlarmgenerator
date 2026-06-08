using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SIM_TIA_DeviceAlarmgenerator.Model;
using SIM_TIA_DeviceAlarmgenerator.Services.DeviceRules;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SIM_TIA_DeviceAlarmgenerator.Services
{
    /// <summary>
    /// Exportiert HMI-Alarmmeldungen in eine Excel-Datei (Sheet "DiscreteAlarms").
    /// - Gerätespezifische Alarme aus <see cref="AlarmDbModel.DeviceTypes"/>
    /// - Instanz-Labels (BMK/NameS1) aus <see cref="AlarmDbModel.DeviceInstanceLabels"/>
    /// - Applikationsalarme aus <see cref="AlarmDbModel.ApplicationAlarms"/> werden zwischen "Omode" und "SQ" eingefügt.
    /// - Optional: appAlarmTextMap (DevType, Bit) → Text überschreibt Gerätestandardtexte.
    /// </summary>
    public sealed class HmiAlarmsExporter
    {
        /// <summary>
        /// Exportiert alle diskreten Alarme in eine Excel-Datei (Tab „DiscreteAlarms“).
        /// 
        /// - Kopfzeile wird analog zum Alt-Tool aufgebaut.
        /// - Geräte/Instanzen/Bit-Positionen werden entsprechend der 16er-Wortgrenzen allokiert.
        /// - Die Alarmtexte werden aus den DeviceRules (ErrorsTemplate) abgeleitet und je Gerätetyp
        ///   exakt wie im Alt-Tool zusammengesetzt:
        ///   * Prefix: laufende Zeilennummer (RowId)
        ///   * Dann Geräteinformation (BMK, NameS1, NameS2) im typ-spezifischen Format
        ///   * Standardtext des Bits (aus ErrorsTemplate)
        ///   * Gerätespezifische Suffixe (z. B. B...GS/AS/MS bei 2DS/2DSPM für bestimmte Bits)
        ///   * Reserve-Texte pro Typ (z. B. "Omode_Reserve_d.{Instanz}", "SQ_Reserve_d.{Instanz}", ...)
        /// - Die Alarmtexte für Alm16 werden aus dem Sheet "AppAlarm (_AA)" entnommen
        /// - Klasse "Warnings" wird ohne Ack-Tag exportiert (wie im Alt-Tool).
        /// - Optional kann <paramref name="appAlarmTextMap"/> bestimmte Bittexte überschreiben
        ///   (Schlüssel: (DevType, BitIndex0Based)); die Klasse bleibt aus der Rule.
        /// </summary>
        /// <param name="outputPath">Zieldatei (xlsx).</param>
        /// <param name="model">Die vollständige Alarmdatenbank (Device-Typen, Labels, App-Alarme).</param>
        /// <param name="appAlarmTextMap">
        ///  Optionales Mapping zum Überschreiben einzelner Bit-Texte je Gerätetyp.
        ///  Schlüssel: (DevType, BitIndex0Based); Wert: der reine Alarmtext (ohne RowId).
        /// </param>
        /// <exception cref="ArgumentNullException">Wenn <paramref name="model"/> null ist.</exception>
        public void Export(
            string outputPath,
            AlarmDbModel model,
            IReadOnlyDictionary<(string DevType, int Bit), string>? appAlarmTextMap = null)
        {
            if (model is null) throw new ArgumentNullException(nameof(model));

            using var wb = new XSSFWorkbook();
            var sh = wb.CreateSheet("DiscreteAlarms");

            // Kontext für Alm16: aktuelle Instanz (1-basiert) für ComposeDeviceText
            int currentDeviceIndex1Based = 1;

            // ------------------------------
            // Kopfzeile analog HMIAlarmsTest.xlsx
            // ------------------------------
            string[] headers =
            {
                "ID",
                "Name",
                "Alarm text [de-DE], Alarm text",
                "FieldInfo [Alarm text]",
                "Class",
                "Trigger tag",
                "Trigger bit",
                "PLC acknowledgement tag",
                "PLC acknowledgement bit",
                "Counter tag",
                "PLC acknowledgement bit.1",
                "Group",
                "Report",
                "Info text [de-DE], Info text"
            };

            var headerRow = sh.CreateRow(0);
            for (int i = 0; i < headers.Length; i++)
                headerRow.CreateCell(i).SetCellValue(headers[i]);

            // Kurz-Handles aus dem Model
            var types = model.DeviceTypes ?? new List<DeviceTypeInfo>();
            var labels = model.DeviceInstanceLabels ??
                         new Dictionary<string, Dictionary<int, DeviceInstanceLabel>>(StringComparer.OrdinalIgnoreCase);
            var appAlm = model.ApplicationAlarms ?? new List<AlarmBit>();

            // Laufende Indizes
            int rowIndex = 1;        // nächste freie Tabellenzeile (≙ RowId im Textprefix)
            int id = 1;              // fortlaufende ID
            int globalBitIndex = 0;  // für die Bit-/Wort-Allokation (pro Gerät)

            // ------------------------------
            // Helpers
            // ------------------------------
            void WriteRow(string name, string text, string @class, string triggerTag, int trigBit, string ackTag, int ackBit)
            {
                var r = sh.CreateRow(rowIndex++);
                int c = 0;

                r.CreateCell(c++).SetCellValue(id++);          // ID
                r.CreateCell(c++).SetCellValue(name);          // Name
                r.CreateCell(c++).SetCellValue(text);          // Alarm text
                r.CreateCell(c++).SetCellValue("");            // FieldInfo
                r.CreateCell(c++).SetCellValue(@class);        // Class
                r.CreateCell(c++).SetCellValue(triggerTag);    // Trigger tag
                r.CreateCell(c++).SetCellValue(trigBit);       // Trigger bit

                // Alt-Tool: Warnings ohne Ack
                if (string.Equals(@class, "Warnings", StringComparison.OrdinalIgnoreCase))
                {
                    r.CreateCell(c++).SetCellValue("<No value>"); // PLC acknowledgement tag
                    r.CreateCell(c++).SetCellValue(0);            // PLC acknowledgement bit
                }
                else
                {
                    r.CreateCell(c++).SetCellValue(ackTag);       // PLC acknowledgement tag
                    r.CreateCell(c++).SetCellValue(ackBit);       // PLC acknowledgement bit
                }

                r.CreateCell(c++).SetCellValue("<No value>"); // Counter tag
                r.CreateCell(c++).SetCellValue(0);            // PLC acknowledgement bit.1
                r.CreateCell(c++).SetCellValue("<No value>"); // Group
                r.CreateCell(c++).SetCellValue(false);        // Report
                r.CreateCell(c++).SetCellValue("<No value>"); // Info text
            }

            /// <summary>
            /// Writes an empty padding row for an unused alarm bit.
            /// </summary>
            /// <param name="triggerTag">The trigger tag name.</param>
            /// <param name="triggerBit">The trigger bit index.</param>
            /// <param name="ackTag">The acknowledgement tag name.</param>
            /// <param name="ackBit">The acknowledgement bit index.</param>
            void WriteEmptyBitRow(string triggerTag, int triggerBit, string ackTag, int ackBit)
            {
                var r = sh.CreateRow(rowIndex++);
                int c = 0;

                r.CreateCell(c++).SetCellValue(id++);          // ID
                r.CreateCell(c++).SetCellValue("");            // Name
                r.CreateCell(c++).SetCellValue("");            // Alarm text
                r.CreateCell(c++).SetCellValue("");            // FieldInfo
                r.CreateCell(c++).SetCellValue("");            // Class
                r.CreateCell(c++).SetCellValue(triggerTag);    // Trigger tag
                r.CreateCell(c++).SetCellValue(triggerBit);    // Trigger bit
                r.CreateCell(c++).SetCellValue(ackTag);        // PLC acknowledgement tag
                r.CreateCell(c++).SetCellValue(ackBit);        // PLC acknowledgement bit
                r.CreateCell(c++).SetCellValue("<No value>");  // Counter tag
                r.CreateCell(c++).SetCellValue(0);             // PLC acknowledgement bit.1
                r.CreateCell(c++).SetCellValue("<No value>");  // Group
                r.CreateCell(c++).SetCellValue(false);         // Report
                r.CreateCell(c++).SetCellValue("<No value>");  // Info text
            }

            static string Safe(string? s) => string.IsNullOrWhiteSpace(s) ? "" : s.Trim();

            // ------------------------------
            // Hauptschleife über alle Gerätetypen
            // ------------------------------
            #region Alarm Rows

            // Main loop over all configured device types.
            for (int tIndex = 0; tIndex < types.Count; tIndex++)
            {
                var type = types[tIndex];

                string devType = type.DevName?.Trim() ?? string.Empty;
                int devQty = Math.Max(0, type.DevQty);
                int almQty = Math.Max(0, type.AlmQty);

                if (string.IsNullOrWhiteSpace(devType))
                    continue;

                if (devQty <= 0 || almQty <= 0)
                    continue;

                // Each device type starts on a clean word boundary.
                globalBitIndex = AlignToNextWord(globalBitIndex);

                // The label map contains the real configured device instances from the Devices sheet.
                labels.TryGetValue(devType, out var instMap);
                int devicesFound = instMap?.Count ?? 0;

                for (int inst = 1; inst <= devQty; inst++)
                {
                    currentDeviceIndex1Based = inst;

                    string bmk = string.Empty;
                    string nameS1 = string.Empty;
                    string nameS2 = string.Empty;

                    if (instMap != null && instMap.TryGetValue(inst, out var info))
                    {
                        bmk = info?.BmkGroup ?? string.Empty;
                        nameS1 = info?.NameS1 ?? string.Empty;
                    }

                    bool isReserve = inst > devicesFound && devicesFound > 0;

                    #region Alarm Bits

                    for (int bit = 0; bit < almQty; bit++)
                    {
                        int hmiBitIndex = ToHmiByteSwappedBitIndex(globalBitIndex);

                        string name = $"{devType}_{bit}_{inst:000}";

                        var (cls, text) = ResolveClassAndText(
                            devType,
                            name,
                            bit,
                            isReserve,
                            rowIndex,
                            bmk,
                            nameS1,
                            nameS2,
                            appAlarmTextMap
                        );

                        WriteRow(
                            name: name,
                            text: text,
                            @class: cls,
                            triggerTag: "Devices",
                            trigBit: hmiBitIndex,
                            ackTag: "Devices_Ack",
                            ackBit: hmiBitIndex
                        );

                        // The logical alarm bit stream continues inside the same device type.
                        globalBitIndex++;
                    }

                    #endregion
                }

                // The next device type starts on a new word.
                globalBitIndex = AlignToNextWord(globalBitIndex);
            }

            #endregion            // Spaltenbreite optimieren (optional)
            for (int i = 0; i < headers.Length; i++)
                sh.AutoSizeColumn(i);

            // Datei schreiben
            var dir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            using var fs = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None);
            wb.Write(fs);

            // Liefert (Klasse, Text) aus Rule & Map; Text wird anschließend exakt wie Alt-Tool formatiert.
            (string Class, string Text) ResolveClassAndText(
                string devType, string name, int bit, bool isReserve, int rowId,
                string bmk, string nameS1, string nameS2,
                IReadOnlyDictionary<(string DevType, int Bit), string>? map)
            {
                // Rule lookup (für StdText & Class)
                string stdText = $"Alarm {bit}";
                string @class = "Errors";

                if (DeviceAlarmRuleFactory.TryGetRule(devType, out var rule) &&
                    bit >= 0 && bit < rule.AlarmsPerDevice)
                {
                    var tpl = rule.ErrorsTemplate[bit];
                    if (!string.IsNullOrWhiteSpace(tpl.Text)) stdText = tpl.Text.Trim();
                    if (!string.IsNullOrWhiteSpace(tpl.Class)) @class = tpl.Class.Trim();
                }

                // appAlarmTextMap überschreibt den inhaltlichen Text (Klasse bleibt aus Rule)
                if (map != null && map.TryGetValue((devType, bit), out var mapped) && !string.IsNullOrWhiteSpace(mapped))
                {
                    // RowId-Prefix nicht vergessen
                    return (@class, $"{rowId} {mapped.Trim()}");
                }

                // Komposition im Alt-Tool-Stil
                string composed = ComposeDeviceText(devType, name, bit, isReserve, rowId, bmk, nameS1, nameS2, stdText);

                // Baut den Alarmtext wie im Alt-Tool – basierend auf DeviceType, Bit, Reserve/Labels und StdText.
                string ComposeDeviceText(
                    string devType, string name, int bit, bool isReserve, int rowId,
                    string bmk, string nameS1, string nameS2, string stdText)
                {
                    bmk = Safe(bmk);
                    nameS1 = Safe(nameS1);
                    nameS2 = Safe(nameS2);


                    // Einige Typen haben besondere "Normal"-Formate (mit/ohne Unterstrich nach RowId)
                    // sowie spezielle Reserve-Texte und Suffixe.
                    string normalFmtUnderscoreRow = $"_{{0}} {{1}} {{2}} {{3}}"; // z.B. BC/CS/RIT_Weiss
                    string normalFmtDefault = $"{{0}}_{{1}} {{2}} {{3}}"; // Standard (BMK _ NameS1 <space> NameS2)
                    string normalFmtOmode = $"{{0}} {{1}} {{2}}";       // OMODE: BMK <space> NameS1 <space> StdText
                    string normalFmtDenso = $"{{0}} {{1}} {{2}}";       // identisch zu vielen anderen

                    string reserve;
                    string text;

                    // Utility: B{BMK ohne die ersten 2 Zeichen}{NameS1}.<SUF>
                    string BuildBmkSuffix(string suffix)
                    {
                        if (string.IsNullOrEmpty(bmk) || bmk.Length < 3) return "";
                        return $" B{bmk.Substring(2)}{nameS1}{suffix}";
                    }

                    var devKey = (devType ?? string.Empty)
                        .Trim()
                        .TrimEnd('_', ' ')
                        .ToUpperInvariant();

                    switch (devKey)
                    {
                        case "OMODE":
                            reserve = $"Omode_Reserve_d.{rowId}";
                            text = isReserve
                                ? $"{reserve} {stdText}"
                                : string.Format(normalFmtOmode, bmk, nameS1, stdText);
                            break;

                        case "ALM16":
                            {
                                reserve = $"Alm16_Reserve_d.{rowId}";

                                // AppAlarm-Zeile wie im Alt-Tool:
                                //   Index = (Instanz-1)*16 + Bit
                                int appIdx = (currentDeviceIndex1Based - 1) * 16 + bit;

                                // Safes Lesen aus der bereits vorhandenen App-Alarm-Liste:
                                string appText = "";
                                if (appIdx >= 0 && appIdx < appAlm.Count)
                                {
                                    var aa = appAlm[appIdx];
                                    // bevorzugt "Name" + „Alarmtext“-Spalte, sonst Name als Fallback
                                    appText = Safe(!string.IsNullOrWhiteSpace(aa.AlarmText) ? $"{bmk}: {aa.AlarmText}" : $"{name} !Alarmtext nicht definiert!");
                                }

                                if (string.IsNullOrEmpty(appText))
                                {
                                    // Fallback identisch zum Alt-Tool: Reserve- oder Standardtext
                                    text = isReserve
                                        ? $"{reserve} {stdText}"
                                        : $"{bmk} {stdText}";
                                }
                                else
                                {
                                    // Alt-Tool-Prinzip: RowId + " " + Text aus AppAlarm (_AA)
                                    // Die RowId wird jetzt weggelassen, da im TIA-Alarmfenster als eigene Spalte vorhanden
                                    text = $"{appText}";
                                }
                                break;
                            }

                        case "SQ":
                            reserve = $"SQ_Reserve_d.{rowId}";
                            text = isReserve
                                ? $"{reserve} {stdText}"
                                : string.Format(normalFmtDefault, bmk, nameS1, nameS2, stdText);
                            break;

                        case "2DS":
                            reserve = $"2_DS Reserver_d.{rowId}";
                            if (isReserve)
                            {
                                text = $"{reserve} {stdText}";
                            }
                            else
                            {
                                string suffix = bit switch
                                {
                                    2 => BuildBmkSuffix(".GS"),
                                    3 => BuildBmkSuffix(".AS"),
                                    _ => ""
                                };
                                text = string.Format(normalFmtDefault, bmk, nameS1, nameS2, stdText) + suffix;
                            }
                            break;

                        case "2DSPM":
                            reserve = $"2_DS Reserver_d.{rowId}";
                            if (isReserve)
                            {
                                text = $"{reserve} {stdText}";
                            }
                            else
                            {
                                string suffix = bit switch
                                {
                                    2 => BuildBmkSuffix(".GS"),
                                    3 => BuildBmkSuffix(".AS"),
                                    4 => BuildBmkSuffix(".MS"),
                                    5 => BuildBmkSuffix(".GS"),
                                    6 => BuildBmkSuffix(".AS"),
                                    _ => ""
                                };
                                text = string.Format(normalFmtDefault, bmk, nameS1, nameS2, stdText) + suffix;
                            }
                            break;

                        case "CONVM":
                            reserve = $"ConvM_Reserve_d.{rowId}";
                            text = isReserve
                                ? $"{reserve} {stdText}"
                                : string.Format(normalFmtDefault, bmk, nameS1, nameS2, stdText);
                            break;

                        case "AI":
                            reserve = $"AI_Reserve .d_.{rowId}";
                            text = isReserve
                                ? $"{reserve} {stdText}"
                                : string.Format(normalFmtDefault, bmk, nameS1, nameS2, stdText);
                            break;

                        case "BC":
                            reserve = $"BC_Reserve_d.{rowId}";
                            text = isReserve
                                ? $"{reserve} {stdText}"
                                : string.Format(normalFmtUnderscoreRow, bmk, nameS1, nameS2, stdText);
                            break;

                        case "CS":
                            reserve = $"´CS_Reserve_d.{rowId}"; // Achtung: Alt-Tool mit Akzent vor CS
                            text = isReserve
                                ? $"{reserve} {stdText}"
                                : string.Format(normalFmtUnderscoreRow, bmk, nameS1, nameS2, stdText);
                            break;

                        case "CAM_XG":
                            reserve = $"Cam_XG_Reserve_d.{rowId}";
                            text = isReserve
                                ? $"{reserve} {stdText}"
                                : string.Format(normalFmtDefault, bmk, nameS1, nameS2, stdText);
                            break;

                        case "VISIONS":
                            reserve = $"VisionS_Reserve_d.{rowId}";
                            text = isReserve
                                ? $"{reserve} {stdText}"
                                : string.Format(normalFmtDenso, bmk, nameS1, stdText);
                            break;

                        case "RFID":
                            reserve = $"RFID_Reserve_d.{rowId}";
                            text = isReserve
                                ? $"{reserve} {stdText}"
                                : string.Format(normalFmtDefault, bmk, nameS1, nameS2, stdText);
                            break;

                        case "COM_RS":
                            reserve = $"COm_RS_Reserve_d.{rowId}";
                            text = isReserve
                                ? $"{reserve} {stdText}"
                                : string.Format(normalFmtDefault, bmk, nameS1, nameS2, stdText);
                            break;

                        case "PN":
                            reserve = $"Reserve PN.{rowId}";
                            text = isReserve
                                ? $"{reserve} {stdText}"
                                : string.Format(normalFmtDefault, bmk, nameS1, nameS2, stdText);
                            break;

                        case "KEYENCEMIC":
                        case "KEYENCE_MC":
                            reserve = $"KeyenceMic_Reserve_d.{rowId}";
                            text = isReserve
                                ? $"{reserve} {stdText}"
                                : string.Format(normalFmtDefault, bmk, nameS1, nameS2, stdText);
                            break;

                        case "DO":
                            reserve = $"DO_Reserve_d.{rowId}";
                            text = isReserve
                                ? $"{reserve} {stdText}"
                                : string.Format(normalFmtDefault, bmk, nameS1, nameS2, stdText);
                            break;

                        case "DI":
                            reserve = $"DI_Reserve_d.{rowId}";
                            text = isReserve
                                ? $"{reserve} {stdText}"
                                : string.Format(normalFmtDefault, bmk, nameS1, nameS2, stdText);
                            break;

                        case "DENSO":
                            reserve = $"Denso_Reserve_d.{rowId}";
                            text = isReserve
                                ? $"{reserve} {stdText}"
                                : string.Format(normalFmtDenso, bmk, nameS1, stdText);
                            break;

                        case "HT_KSV":
                            reserve = $"HT_KSV_Reserve_d.{rowId}";
                            text = isReserve
                                ? $"{reserve} {stdText}"
                                : string.Format(normalFmtDefault, bmk, nameS1, nameS2, stdText);
                            break;

                        case "APOSI":
                            reserve = $"Aposi_Reserve_d.{rowId}";
                            text = isReserve
                                ? $"{reserve} {stdText}"
                                : string.Format(normalFmtDefault, bmk, nameS1, nameS2, stdText);
                            break;

                        case "M_TC":
                            reserve = $"Reserve MTC _d.{rowId}";
                            text = isReserve
                                ? $"{reserve} {stdText}"
                                : string.Format(normalFmtDefault, bmk, nameS1, nameS2, stdText);
                            break;

                        case "LOAD_DIS":
                        case "Load&LOAD&DISPLACEM":
                        case "LOAD&DISPL":
                            reserve = $"Load&Displacem_Reserve_d.{rowId}";
                            text = isReserve
                                ? $"{reserve} {stdText}"
                                : string.Format(normalFmtDefault, bmk, nameS1, nameS2, stdText);
                            break;

                        case "RIT_WEISS":
                            reserve = $"Rit_Weiss_Reserve_d.{rowId}";
                            text = isReserve
                                ? $"{reserve} {stdText}"
                                : string.Format(normalFmtUnderscoreRow, bmk, nameS1, nameS2, stdText);
                            break;

                        default:
                            // Generisches Fallback (wie vormals): RowId + "BMK NameS1 DevType Fehler <bit>"
                            var parts = new[] { bmk, nameS1, devType }
                                .Where(s => !string.IsNullOrWhiteSpace(s));
                            string baseText = string.Join(" ", parts);
                            if (string.IsNullOrWhiteSpace(baseText)) baseText = devType;
                            text = $"{baseText} Fehler {bit}";
                            break;
                    }

                    return text;
                }

                return (@class, composed);
            }

        }

        #region HMI Bit Mapping

        /// <summary>
        /// Converts an absolute logical alarm bit index to the HMI bit index with swapped low and high byte inside each 16-bit word.
        /// </summary>
        /// <param name="logicalBitIndex">The absolute logical alarm bit index.</param>
        /// <returns>The HMI bit index with swapped byte order inside the related 16-bit word.</returns>
        private static int ToHmiByteSwappedBitIndex(int logicalBitIndex)
        {
            if (logicalBitIndex < 0)
                throw new ArgumentOutOfRangeException(nameof(logicalBitIndex), "The logical bit index cannot be negative.");

            int wordStartBitIndex = (logicalBitIndex / 16) * 16;
            int bitIndexInsideWord = logicalBitIndex % 16;

            // The HMI import expects low byte and high byte to be swapped inside each 16-bit word.
            int swappedBitIndexInsideWord = bitIndexInsideWord < 8
                ? bitIndexInsideWord + 8
                : bitIndexInsideWord - 8;

            return wordStartBitIndex + swappedBitIndexInsideWord;
        }

        /// <summary>
        /// Aligns a logical bit index to the first bit of the next 16-bit word if needed.
        /// </summary>
        /// <param name="logicalBitIndex">The logical bit index to align.</param>
        /// <returns>The aligned logical bit index.</returns>
        private static int AlignToNextWord(int logicalBitIndex)
        {
            if (logicalBitIndex < 0)
                throw new ArgumentOutOfRangeException(nameof(logicalBitIndex), "The logical bit index cannot be negative.");

            int remainder = logicalBitIndex % 16;

            if (remainder == 0)
                return logicalBitIndex;

            return logicalBitIndex + (16 - remainder);
        }

        #endregion
    }
}
