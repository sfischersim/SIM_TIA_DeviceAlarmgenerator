using SIM_TIA_DeviceAlarmgenerator.Model;
using System;
using System.Collections.Generic;

namespace SIM_TIA_DeviceAlarmgenerator.Services
{
    /// <summary>
    /// Generischer Alarm-Builder, der über eine <see cref="IDeviceAlarmRule"/> konfiguriert wird
    /// und aus Device-/Excel-Daten <see cref="AlarmRow"/>-Einträge erzeugt.
    /// </summary>
    /// <remarks>
    /// Der Builder ist zustandsarm; die Logik (Tag-Namen, Texte, Anzahl Alarme je Gerät)
    /// steckt in der jeweils übergebenen Regel (<see cref="IDeviceAlarmRule"/>).
    /// </remarks>
    public sealed class GenericDeviceAlarmBuilder
    {
        /// <summary>
        /// Regel/Strategie, die Tag-Namen, Textbausteine und Alarmanzahl vorgibt.
        /// </summary>
        private readonly IDeviceAlarmRule _rule;

        /// <summary>
        /// Initialisiert einen generischen Builder für einen konkreten Device-Typ.
        /// </summary>
        /// <param name="rule">Regel/Strategie für den Device-Typ.</param>
        /// <exception cref="ArgumentNullException"><paramref name="rule"/> ist <see langword="null"/>.</exception>
        public GenericDeviceAlarmBuilder(IDeviceAlarmRule rule)
        {
            if (rule == null)
                throw new ArgumentNullException(nameof(rule));
            _rule = rule;
        }

        /// <summary>
        /// Erzeugt die Alarmzeilen für den konfigurierten Device-Typ aus der „Devices“-Matrix und einer Geräteanzahl.
        /// </summary>
        /// <param name="devicesMatrix">Matrix des Sheets „Devices“ im Format <c>[Zeile, Spalte]</c>.</param>
        /// <param name="quantity">Geräteanzahl aus „StA-Cfg“ (≤0 erzeugt keine Ausgaben).</param>
        /// <returns>Sequenz erzeugter <see cref="AlarmRow"/> im Ziel-Layout (14 Spalten).</returns>
        /// <exception cref="ArgumentNullException"><paramref name="devicesMatrix"/> ist <see langword="null"/>.</exception>
        /// <remarks>
        /// - Die globale Zeilennummer beginnt bei 1 und wird in <see cref="AlarmRow.RowNo"/> übernommen.<br/>
        /// - Der Meldungstext wird über die Regel (<c>ComposeMessage</c>) erstellt; bei Bedarf werden
        ///   Spalten aus der Devices-Matrix gelesen (siehe <see cref="ComposeDefaultMessage"/>).<br/>
        /// - <see cref="AlarmRow.RefIndex"/> wird nach bekannter Formel aus der laufenden Zeile berechnet.
        /// </remarks>
        public IEnumerable<AlarmRow> Build(string[,] devicesMatrix, int quantity)
        {
            if (devicesMatrix is null) throw new ArgumentNullException(nameof(devicesMatrix));
            if (quantity <= 0) yield break;

            // Block suchen (Start/End analog zur Alt-Logik).
            var (startRow1, endRow1) = FindDeviceBlock(devicesMatrix, _rule.DeviceTypeName);
            int span = Math.Max(0, endRow1 - startRow1);   // inkl. Kopf/Leer
            int elementsFromTab = Math.Max(0, span - 2);   // -2 für Kopf/Trenner

            int row = 1; // globale Zeilennummer (1-basiert)

            for (int devIdx = 1; devIdx <= quantity; devIdx++)
            {
                for (int alarmIdx = 0; alarmIdx < _rule.AlarmsPerDevice; alarmIdx++)
                {
                    var (txt, cls) = SafeTemplate(_rule.ErrorsTemplate, alarmIdx);
                    string @class = string.IsNullOrWhiteSpace(cls) ? "Errors" : cls;

                    var ctx = new DeviceMessageContext(
                        globalRowNo1Based: row,
                        deviceIndex1Based: devIdx,
                        alarmIndex0Based: alarmIdx,
                        devicesMatrix: devicesMatrix,
                        dataStartRow1Based: startRow1,
                        blockSpan: span,
                        elementsFromTab: elementsFromTab
                    );

                    var r = new AlarmRow
                    {
                        RowNo = row.ToString(),
                        TagName = _rule.ComposeTagName(alarmIdx, devIdx),
                        Class = @class,
                        Message = _rule.ComposeMessage(ctx)
                    };

                    // RefIndex-Berechnung identisch zur bekannten Formel
                    int num4 = (row - 1) / 8;
                    int mod8 = (row - 1) % 8;
                    int refIndex = (num4 % 2 == 1) ? num4 * 8 + mod8 - 8 : num4 * 8 + mod8 + 8;
                    r.RefIndex = refIndex.ToString();

                    // Ack-Bereich (Errors → Devices_Ack/RefIndex, sonst leer)
                    if (r.Class.Equals("Errors", StringComparison.OrdinalIgnoreCase))
                    {
                        r.AckArea = "Devices_Ack";
                        r.AckIndex = r.RefIndex;
                    }
                    else
                    {
                        r.AckArea = "<No value>";
                        r.AckIndex = "0";
                    }

                    yield return r;
                    row++;
                }
            }
        }

        /// <summary>
        /// Liefert <paramref name="matrix"/>[<paramref name="r1"/>, <paramref name="c"/>] sicher zurück
        /// (leer bei Out-of-Range). <paramref name="r1"/> ist 1-basiert.
        /// </summary>
        /// <param name="matrix">Quellmatrix.</param>
        /// <param name="r1">1-basierter Zeilenindex.</param>
        /// <param name="c">0-basierter Spaltenindex.</param>
        /// <returns>Zelleninhalt oder leerer String, wenn außerhalb des gültigen Bereichs.</returns>
        private static string Safe(string[,] matrix, int r1, int c)
        {
            int r = r1 - 1;
            if (r < 0 || c < 0) return string.Empty;
            if (r >= matrix.GetLength(0) || c >= matrix.GetLength(1)) return string.Empty;
            return matrix[r, c] ?? string.Empty;
        }

        /// <summary>
        /// Sucht den Datenblock eines Device-Typs in der „Devices“-Matrix.
        /// </summary>
        /// <param name="matrix">Quellmatrix.</param>
        /// <param name="deviceName">Gerätetypbezeichnung/Blocküberschrift in Spalte 0.</param>
        /// <returns>
        /// Tupel aus 1-basiertem Start- und Endindex (<c>startRow1</c>, <c>endRow1</c>).
        /// Gibt (0,0) zurück, wenn der Block nicht gefunden wird.
        /// </returns>
        private static (int startRow1, int endRow1) FindDeviceBlock(string[,] matrix, string deviceName)
        {
            int rows = matrix.GetLength(0);

            // Header finden
            int header1 = 0;
            for (int r = 0; r < rows; r++)
            {
                string c0 = matrix[r, 0] ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(c0) &&
                    c0.Equals(deviceName, StringComparison.CurrentCultureIgnoreCase))
                {
                    header1 = r + 1;
                    break;
                }
            }
            if (header1 == 0) return (0, 0);

            // Ende: erste Leerzeile in Spalte 0 nach Header
            int end1 = header1;
            for (int r = header1; r < rows; r++)
            {
                string c0 = matrix[r, 0] ?? string.Empty;
                if (string.IsNullOrWhiteSpace(c0))
                {
                    end1 = r;
                    break;
                }
                end1 = r + 1;
            }

            // Daten starten eine Zeile nach dem Header
            int startData1 = header1 + 1;
            return (startData1, end1);
        }

        /// <summary>
        /// Liefert (Text, Class) aus der Vorlage index-sicher; Standard: leerer Text, Klasse „Errors“.
        /// </summary>
        /// <param name="tpl">Fehlervorlage der Regel.</param>
        /// <param name="index">0-basierter Alarmindex.</param>
        /// <returns>Gewähltes Template-Paar oder Fallback (leer, „Errors“).</returns>
        private static (string Text, string Class) SafeTemplate((string Text, string Class)[] tpl, int index)
        {
            if (tpl is null || index < 0 || index >= tpl.Length) return (string.Empty, "Errors");
            var (t, c) = tpl[index];
            return (t ?? string.Empty, string.IsNullOrWhiteSpace(c) ? "Errors" : c);
        }

        /// <summary>
        /// Standard-Textaufbau für Regeln: liest Spalten 1/2 aus der „Devices“-Matrix,
        /// andernfalls wird ein Reserve-Präfix mit der Instanznummer verwendet.
        /// </summary>
        /// <param name="ctx">Kontext der aktuellen Meldungserzeugung.</param>
        /// <param name="reservePrefix">Reserve-Präfix (z. B. Gerätetyp) für den Fallback.</param>
        /// <param name="errorText">Fehler-/Warnungstext aus dem Template.</param>
        /// <returns>Aufbereiteter Meldungstext inklusive globaler Zeilennummer.</returns>
        public static string ComposeDefaultMessage(DeviceMessageContext ctx, string reservePrefix, string errorText)
        {
            // Bedingung analog Alt-Logik: (deviceIdx < span - 1 && span > 0 && elementsFromTab > 0)
            if (ctx.DeviceIndex1Based < (ctx.BlockSpan - 1) && ctx.BlockSpan > 0 && ctx.ElementsFromTab > 0)
            {
                string col1 = Safe(ctx.DevicesMatrix, ctx.DataStartRow1Based + ctx.DeviceIndex1Based, 1);
                string col2 = Safe(ctx.DevicesMatrix, ctx.DataStartRow1Based + ctx.DeviceIndex1Based, 2);
                return $"{ctx.GlobalRowNo1Based} {col1} {col2} {errorText}".Trim();
            }
            return $"{ctx.GlobalRowNo1Based} {reservePrefix}{ctx.DeviceIndex1Based} {errorText}".Trim();
        }
    }
}
