using SIM_TIA_DeviceAlarmgenerator.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SIM_TIA_DeviceAlarmgenerator.Services
{
    /// <summary>
    /// Erzeugt den Siemens-DB-Quelltext im „STRUCT + BEGIN“-Layout auf Basis eines <see cref="AlarmDbModel"/>.
    /// </summary>
    /// <remarks>
    /// Der Builder formatiert Header, STRUCT-Bereich und BEGIN-Initialisierungen TIA-kompatibel.
    /// Er arbeitet zustandslos und kann gefahrlos mehrfach verwendet werden.
    /// </remarks>
    public static class AlarmDbTextBuilder
    {
        /// <summary>
        /// Baut den vollständigen DB-Quelltext (TIA-kompatibles .db-Format) aus dem übergebenen Modell.
        /// </summary>
        /// <param name="m">Datenmodell mit Header, Standard-/Applikationsalarmen, Geräteeinstellungen und Maxima.</param>
        /// <returns>Den generierten DB-Quelltext als String.</returns>
        /// <remarks>
        /// Erwartet ein vollständig initialisiertes Modell. Leere Listen werden als Platzhalter kommentiert,
        /// negative oder 0-Werte werden auf sinnvolle Minimalwerte normalisiert (z. B. 1 Zeile/Wort).
        /// </remarks>
        public static string Build(AlarmDbModel m)
        {
            const string NL = "\r\n";
            var sb = new StringBuilder(64 * 1024);

            // Header
            sb.Append("DATA_BLOCK ").Append('"').Append(m.Name).Append('"').Append(NL)
              .Append("TITLE = ").Append(m.Title).Append(NL)
              .Append("{ S7_Optimized_Access := '").Append(m.OptimizedAccess ? "TRUE" : "FALSE").Append("' }").Append(NL)
              .Append("AUTHOR : '").Append(m.Author).Append('\'').Append(NL)
              .Append("FAMILY : ").Append(m.Family).Append(NL)
              .Append("VERSION : ").Append(m.Version).Append(NL);
            if (m.NonRetain) sb.Append("NON_RETAIN").Append(NL);

            // Kopf-Kommentar
            sb.Append("//Standardschnittstelle für alle Standard, Device und Applikationsalarme zum HMI.").Append(NL)
              .Append("// inklusive der Konfiguration der verwnedeten Device mit Anzahl und Anzahl Meldung pro Devicetyp.").Append(NL)
              .Append("// !!! Diesen Baustein nicht manuell ändern bzw. Änderung in dem dazugehörigen Excel Projektierungstool nachpflegen.").Append(NL);

            // STRUCT
            sb.Append("STRUCT ").Append(NL);

            // Std
            sb.Append("      Std : Struct   // Standard Alarme").Append(NL);
            if (m.StandardAlarms?.Count > 0)
            {
                foreach (var bit in m.StandardAlarms)
                {
                    var name = string.IsNullOrWhiteSpace(bit.Name) ? "Unnamed" : bit.Name.Trim();
                    sb.Append("         ").Append(name).Append(" : Bool;");
                    if (!string.IsNullOrWhiteSpace(bit.AlarmText)) sb.Append("   // ").Append(bit.AlarmText.Trim());
                    sb.Append(NL);
                }
            }
            else
            {
                sb.Append("         // (keine StandardAlarms-Einträge vorhanden)").Append(NL);
            }
            sb.Append("END_STRUCT;").Append(NL);

            // Std_Ack
            var ackCount = m.StdAckReserveCount <= 0 ? 18 : m.StdAckReserveCount;
            sb.Append("      Std_Ack : Struct   // Standard Alarme Quittierbereich").Append(NL);
            for (int i = 1; i <= ackCount; i++)
                sb.Append("         reserv_").Append(i).Append(" : Bool;").Append(NL);
            sb.Append("END_STRUCT;").Append(NL);

            // Arrays + Maxima
            var rows = m.DeviceSettingRows <= 0 ? 1 : m.DeviceSettingRows;
            sb.Append($"      DeviceSetting : Array[1..{rows}, 0..2] of UInt;   // setting für collectAlarm").Append(NL);
            sb.Append("      DeviceSettingMax : Int;   // UBound DevSetting").Append(NL);

            var words = m.DevicesWordCount <= 0 ? 1 : m.DevicesWordCount;
            sb.Append($"      Devices : Array[0..{words - 1}] of Word;   // Device Alarme").Append(NL);
            sb.Append($"      Devices_Ack : Array[0..{words - 1}] of Word;   // Device Alarme Quittierbereich").Append(NL);
            sb.Append("      DevicesMax : Int;   // UBound AlarmMax").Append(NL);

            var errCnt = m.ErrNoBufferCount <= 0 ? 10 : m.ErrNoBufferCount;
            sb.Append($"      ErrNoBuffer : Array[0..{errCnt - 1}] of UInt;   // ersten aktive Störungen Nummerisch [0]=0 dann keiner aktiv").Append(NL);
            sb.Append("      ErrNoBufferMax : Int;   // UBound ErrNoBuffer").Append(NL);

            sb.Append("      Std_Val : Struct   // Variablen zu Standardalarmen").Append(NL)
              .Append("         HMI_AL_DevTyp : Int;   // DevTyp mit Fehler").Append(NL)
              .Append("         PCS_TA_No : UInt;   // PCS Transktionsnummer").Append(NL)
              .Append("         PD_Ctrl_Idx : Int;   // Teile Daten Manger Idx").Append(NL)
              .Append("   END_STRUCT;").Append(NL);

            sb.Append("END_STRUCT;").Append(NL).Append(NL);

            // BEGIN (Initialisierungen)
            sb.Append(" BEGIN").Append(NL);

            // DeviceSetting initialisieren (1..rows)
            if (m.DeviceSetting != null)
            {
                for (int i = 0; i < rows; i++)
                {
                    var q = m.DeviceSetting[i, 0];
                    var a = m.DeviceSetting[i, 1];
                    var e = m.DeviceSetting[i, 2];
                    if (q != 0) sb.Append($"DeviceSetting[{i + 1}, 0] := {q};").Append(NL);
                    else sb.Append($"DeviceSetting[{i + 1}, 0] := 0;").Append(NL);
                    
                    if (a != 0) sb.Append($"DeviceSetting[{i + 1}, 1] := {a};").Append(NL);
                    else sb.Append($"DeviceSetting[{i + 1}, 1] := 0;").Append(NL);

                    if (e != 0) sb.Append($"DeviceSetting[{i + 1}, 2] := {e};").Append(NL);
                    else sb.Append($"DeviceSetting[{i + 1}, 2] := 0;").Append(NL);

                }
            }

            sb.Append($"DeviceSettingMax := {m.DeviceSettingMax};").Append(NL);
            sb.Append($"DevicesMax := {m.DevicesMax};").Append(NL);
            sb.Append($"ErrNoBufferMax := {m.ErrNoBufferMax};").Append(NL);

            sb.Append("END_DATA_BLOCK").Append(NL);
            return sb.ToString();
        }
    }
}
