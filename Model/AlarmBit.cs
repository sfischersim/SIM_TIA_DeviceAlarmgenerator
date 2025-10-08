using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SIM_TIA_DeviceAlarmgenerator.Model
{
    /// <summary>
    /// Repräsentiert ein einzelnes Bool-Alarmbit innerhalb eines Gerätes oder Moduls.
    /// </summary>
    /// <remarks>
    /// Enthält den Signalnamen sowie optional einen Kommentar, wie er z. B. aus einer Projekttabelle,
    /// Excel-Quelle oder HMI-Textbeschreibung stammen kann.
    /// </remarks>
    public class AlarmBit
    {
        /// <summary>
        /// Bezeichnung des Alarmbits, z. B. der Symbol- oder Variablenname.
        /// </summary>
        /// <remarks>
        /// Wird typischerweise für die Generierung der Alarmtexte oder für die Zuordnung im TIA-Projekt verwendet.
        /// </remarks>
        public string Name { get; set; } = "";
        /// <summary>
        /// Optionaler Kommentar zum Alarmbit, z. B. aus einer Quelltabelle oder als Beschreibung für das HMI.
        /// </summary>
        /// <remarks>
        /// Kann <see langword="null"/> sein, wenn kein Kommentar angegeben wurde.
        /// </remarks>
        public string? Comment { get; set; }
    }
}
