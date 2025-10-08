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
        private enum InsertMode { BeforeSq, AfterOmode, BeforeFirstSq, AfterLast }

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
        /// - Klasse "Warnings" wird ohne Ack-Tag exportiert (wie im Alt-Tool).
        /// - Application-Alarme (falls vorhanden) werden abhängig von der Position von Omode/SQ eingefügt.
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

            // ------------------------------
            // Einfüge-Strategie für Applikationsalarme bestimmen
            // ------------------------------
            int idxOmode = types.FindIndex(t => string.Equals(t.DevName, "Omode", StringComparison.OrdinalIgnoreCase));
            int idxSq = types.FindIndex(t => string.Equals(t.DevName, "SQ", StringComparison.OrdinalIgnoreCase));

            InsertMode insertion;
            if (idxOmode >= 0 && idxSq >= 0 && idxOmode < idxSq)
                insertion = InsertMode.BeforeSq;     // Omode vor SQ → App-Alarme dazwischen
            else if (idxOmode >= 0 && idxSq < 0)
                insertion = InsertMode.AfterOmode;   // Nur Omode → direkt danach
            else if (idxOmode < 0 && idxSq >= 0)
                insertion = InsertMode.BeforeFirstSq; // Nur SQ → davor
            else
                insertion = InsertMode.AfterLast;    // Weder Omode noch SQ → ans Ende

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

            // App-Alarme (einmalig) ausgeben
            bool appInserted = false;
            void EmitAppAlarmsIfNeeded()
            {
                if (appInserted) return;
                if (appAlm.Count == 0) { appInserted = true; return; }

                int appBit = 0;
                for (int i = 0; i < appAlm.Count; i++)
                {
                    var aa = appAlm[i];
                    string aaName = string.IsNullOrWhiteSpace(aa.Name) ? $"APP_{i:000}" : aa.Name.Trim();
                    string aaText = !string.IsNullOrWhiteSpace(aa.Comment) ? aa.Comment!.Trim() : aaName;

                    // RowId-Prefix wie im Alt-Tool
                    string textFinal = $"{rowIndex} {aaText}";

                    WriteRow(
                        name: aaName,
                        text: textFinal,
                        @class: "Errors", // Alt-Tool: Class kommt aus den App-Daten; hier standardmäßig "Errors"
                        triggerTag: "AppAlarm",
                        trigBit: appBit,
                        ackTag: "AppAlarm_Ack",
                        ackBit: appBit
                    );
                    appBit++;
                }

                appInserted = true;
            }

            static string Safe(string? s) => string.IsNullOrWhiteSpace(s) ? "" : s.Trim();

            // Baut den Alarmtext wie im Alt-Tool – basierend auf DeviceType, Bit, Reserve/Labels und StdText.
            string ComposeDeviceText(
                string devType, int bit, bool isReserve, int rowId,
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

                switch (devType?.Trim())
                {
                    case "Omode":
                        reserve = $"Omode_Reserve_d.{rowId.ToString().PadLeft(0)}";
                        text = isReserve
                            ? $"{reserve} {stdText}"
                            : string.Format(normalFmtOmode, bmk, nameS1, stdText);
                        break;

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

                    case "ConvM":
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

                    case "Cam_XG":
                        reserve = $"Cam_XG_Reserve_d.{rowId}";
                        text = isReserve
                            ? $"{reserve} {stdText}"
                            : string.Format(normalFmtDefault, bmk, nameS1, nameS2, stdText);
                        break;

                    case "VisionS":
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

                    case "Com_RS":
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

                    case "KeyenceMIC":
                    case "Keyence_MC":
                    case "KeyenceMic":
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

                    case "Denso":
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

                    case "APosi":
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

                    case "Load_Dis":
                    case "Load&Displacem":
                    case "Load&Displ":
                        reserve = $"Load&Displacem_Reserve_d.{rowId}";
                        text = isReserve
                            ? $"{reserve} {stdText}"
                            : string.Format(normalFmtDefault, bmk, nameS1, nameS2, stdText);
                        break;

                    case "RIT_Weiss":
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

            // Liefert (Klasse, Text) aus Rule & Map; Text wird anschließend exakt wie Alt-Tool formatiert.
            (string Class, string Text) ResolveClassAndText(
                string devType, int bit, bool isReserve, int rowId,
                string bmk, string nameS1, string nameS2,
                IReadOnlyDictionary<(string DevType, int Bit), string>? map)
            {
                // Rule lookup (für StdText & Class)
                string stdText = $"Fehler {bit}";
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
                string composed = ComposeDeviceText(devType, bit, isReserve, rowId, bmk, nameS1, nameS2, stdText);
                return (@class, composed);
            }

            // ------------------------------
            // Hauptschleife über alle Gerätetypen
            // ------------------------------
            for (int tIndex = 0; tIndex < types.Count; tIndex++)
            {
                var t = types[tIndex];
                string devType = t.DevName?.Trim() ?? "";
                int devQty = Math.Max(0, t.DevQty);
                int almQty = Math.Max(0, t.AlmQty);

                // Vor SQ App-Alarme einfügen (Zwischen-Fall)
                if (!appInserted && insertion == InsertMode.BeforeSq &&
                    string.Equals(devType, "SQ", StringComparison.OrdinalIgnoreCase))
                {
                    EmitAppAlarmsIfNeeded();
                }

                // Gerätespezifische Alarme ausgeben
                if (!string.IsNullOrEmpty(devType) && devQty > 0 && almQty > 0)
                {
                    int wordsPerDevice = (int)Math.Ceiling(almQty / 16.0);
                    int bitsPerDevice = wordsPerDevice * 16;

                    // Wie viele Instanzen sind im Label-Set tatsächlich vorhanden?
                    labels.TryGetValue(devType, out var instMap);
                    int devicesFound = instMap?.Count ?? 0;

                    for (int inst = 1; inst <= devQty; inst++)
                    {
                        int startBit = globalBitIndex;

                        // Instanzlabels (BMK/NameS1/NameS2)
                        string bmk = "";
                        string nameS1 = "";
                        string nameS2 = "";
                        if (instMap != null && instMap.TryGetValue(inst, out var info))
                        {
                            bmk = info?.BmkGroup ?? "";
                            nameS1 = info?.NameS1 ?? "";
                        }

                        bool isReserve = inst > devicesFound && devicesFound > 0;

                        for (int bit = 0; bit < almQty; bit++)
                        {
                            int trigBit = (startBit + bit) % 16;

                            // Name exakt wie früher: <DevType>_<Bit>_<Instanz:000>
                            string name = $"{devType}_{bit}_{inst:000}";

                            // Text & Klasse bestimmen (Rule + Alt-Tool-Format)
                            var (cls, text) = ResolveClassAndText(
                                devType, bit, isReserve, rowIndex,
                                bmk, nameS1, nameS2,
                                appAlarmTextMap
                            );

                            WriteRow(
                                name: name,
                                text: text,
                                @class: cls,
                                triggerTag: "Devices",
                                trigBit: trigBit,
                                ackTag: "Devices_Ack",
                                ackBit: trigBit
                            );
                        }

                        // Nächste Instanz auf Wortgrenze schieben
                        globalBitIndex += bitsPerDevice;
                    }
                }

                // Nach Omode einfügen (wenn SQ fehlt)
                if (!appInserted && insertion == InsertMode.AfterOmode &&
                    string.Equals(devType, "Omode", StringComparison.OrdinalIgnoreCase))
                {
                    EmitAppAlarmsIfNeeded();
                }
            }

            // Falls noch nicht eingefügt:
            if (!appInserted)
            {
                if (insertion == InsertMode.BeforeFirstSq)
                {
                    // Nur SQ vorhanden → davor einfügen
                    EmitAppAlarmsIfNeeded();
                }
                else
                {
                    // Keiner vorhanden → ans Ende
                    EmitAppAlarmsIfNeeded();
                }
            }

            // Spaltenbreite optimieren (optional)
            for (int i = 0; i < headers.Length; i++)
                sh.AutoSizeColumn(i);

            // Datei schreiben
            var dir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            using var fs = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None);
            wb.Write(fs);
        }
    }
}
