using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SIM_TIA_DeviceAlarmgenerator.Model
{
    /// <summary>
    /// Repräsentiert die Label-Information einer Geräteinstanz aus dem Gerätesheet („Devices“).
    /// </summary>
    /// <remarks>
    /// Diese Klasse enthält die Zuordnung zwischen der BMK-Gruppe und dem Namen der Geräteinstanz.
    /// Sie wird beim Generieren der Alarmtexte zur Identifikation einzelner Geräteinstanzen verwendet.
    /// </remarks>
    public sealed class DeviceInstanceLabel
    {
        /// <summary>
        /// BMK-Gruppe oder übergeordnete Kennzeichnung der Geräteinstanz (z. B. Anlagenabschnitt oder Station).
        /// </summary>
        /// <remarks>
        /// Wird genutzt, um Geräteinstanzen logisch zu gruppieren und in der Alarmstruktur zu ordnen.
        /// </remarks>
        public string BmkGroup { get; set; } = "";

        /// <summary>
        /// Name der Geräteinstanz (S1-Ebene), wie er in der Gerätekonfiguration oder im Excel-Sheet angegeben ist.
        /// </summary>
        /// <remarks>
        /// Dient als Referenzbezeichnung innerhalb der generierten Alarm- oder DB-Struktur.
        /// </remarks>
        public string NameS1 { get; set; } = "";
    }
}
