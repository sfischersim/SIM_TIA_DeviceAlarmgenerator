using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SIM_TIA_DeviceAlarmgenerator.Model
{
    /// <summary>
    /// Repräsentiert eine generierte Alarmzeile, die später als Excel-Zeile exportiert wird.
    /// </summary>
    /// <remarks>
    /// Die Spalten sind so angeordnet, dass sie dem Ziel-Layout der HMI-Meldetabelle entsprechen.
    /// Jede Instanz dieser Klasse entspricht einer einzelnen Alarmdefinition (14 Spalten).
    /// </remarks>
    public sealed class AlarmRow
    {
        /// <summary>
        /// Spalte [0]: Zeilennummer der Alarmdefinition in der Excel-Tabelle.
        /// </summary>
        public string RowNo { get; set; } = string.Empty;

        /// <summary>
        /// Spalte [1]: Tag-Name oder Signalbezeichnung, auf die sich der Alarm bezieht.
        /// </summary>
        public string TagName { get; set; } = string.Empty;

        /// <summary>
        /// Spalte [2]: Alarmtext oder Meldungsinhalt, der später auf der HMI angezeigt wird.
        /// </summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>
        /// Spalte [3]: Reservierte Spalte (derzeit ohne funktionale Bedeutung).
        /// </summary>
        public string Reserved3 { get; set; } = string.Empty;

        /// <summary>
        /// Spalte [4]: Klassifizierung des Alarms, z. B. „Errors“ oder „Warnings“.
        /// </summary>
        public string Class { get; set; } = "Errors";

        /// <summary>
        /// Spalte [5]: Kategorie des Alarms, z. B. „Devices“ oder „Standard“.
        /// </summary>
        public string Category { get; set; } = "Devices";

        /// <summary>
        /// Spalte [6]: Referenzindex für Querverweise innerhalb der Alarmstruktur.
        /// </summary>
        public string RefIndex { get; set; } = "0";

        /// <summary>
        /// Spalte [7]: Bereichsbezeichnung für Quittierungen.
        /// </summary>
        public string AckArea { get; set; } = "<No value>";

        /// <summary>
        /// Spalte [8]: Indexwert der Quittierungsstelle.
        /// </summary>
        public string AckIndex { get; set; } = "0";

        /// <summary>
        /// Spalte [9]: Reservierte Spalte (optional für zukünftige Erweiterungen).
        /// </summary>
        public string Col9 { get; set; } = "<No value>";

        /// <summary>
        /// Spalte [10]: Numerischer Wert oder Zähler für erweiterte Alarmfunktionen.
        /// </summary>
        public string Col10 { get; set; } = "0";

        /// <summary>
        /// Spalte [11]: Reservierte Spalte (aktuell nicht verwendet).
        /// </summary>
        public string Col11 { get; set; } = "<No value>";

        /// <summary>
        /// Spalte [12]: Boolescher Wert zur Kennzeichnung von Zuständen.
        /// </summary>
        public string Col12 { get; set; } = "False";

        /// <summary>
        /// Spalte [13]: Reservierte Spalte (aktuell ohne Funktion, Platzhalter für zukünftige Nutzung).
        /// </summary>
        public string Col13 { get; set; } = "<No value>";
    }
}
