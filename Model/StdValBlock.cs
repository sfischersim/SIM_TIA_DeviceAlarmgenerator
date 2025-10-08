using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SIM_TIA_DeviceAlarmgenerator.Model
{
    /// <summary>
    /// Repräsentiert den Standardwert-Block (Std_Val) des Alarmdatenbausteins.
    /// </summary>
    /// <remarks>
    /// Dieser Block enthält Platzhalter- oder Defaultwerte, die im generierten Datenbaustein
    /// im Headerbereich angelegt werden. Er dient als Strukturrahmen ohne Initialwerte.
    /// </remarks>
    public class StdValBlock
    {
        /// <summary>
        /// Kennung des Gerätetyps für HMI-Alarme.
        /// </summary>
        /// <remarks>
        /// Wird im Standard-Alarmdatenbaustein verwendet, um den Typ des Geräts zu identifizieren,
        /// dessen Alarme im HMI sichtbar sind.
        /// </remarks>
        public int HMI_AL_DevTyp { get; set; } = 0;

        /// <summary>
        /// Prozess- oder Teilanlagen-Nummer (PCS_TA_No) für die Zuordnung innerhalb der Steuerung.
        /// </summary>
        /// <remarks>
        /// Wird genutzt, um Alarme und Prozessdaten einer Teilanlage eindeutig zu referenzieren.
        /// </remarks>
        public uint PCS_TA_No { get; set; } = 0;

        /// <summary>
        /// Indexnummer des zugehörigen Prozessdatensatzes oder Steuerungsobjekts.
        /// </summary>
        /// <remarks>
        /// Wird bei der Zuordnung von Alarmelementen zu Prozesssteuerungsindizes verwendet.
        /// </remarks>
        public int PD_Ctrl_Idx { get; set; } = 0;
    }
}
