using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SIM_TIA_DeviceAlarmgenerator.Model
{
    /// <summary>
    /// Repräsentiert das Modell eines generierten Alarm-Datenbausteins (_Alarms) im TIA Portal.
    /// </summary>
    /// <remarks>
    /// Enthält alle Headerinformationen, Standard- und Applikationsalarme sowie Geräteeinstellungen.
    /// Dient als zentrale Struktur für die automatische Generierung des Alarm-Datenbausteins
    /// durch das Device-Alarmgenerator-Tool.
    /// </remarks>
    public class AlarmDbModel
    {
        // Header
        /// <summary>
        /// Name des Alarm-Datenbausteins.
        /// </summary>
        public string Name { get; set; } = "_Alarms";

        /// <summary>
        /// Titel des Alarm-Datenbausteins, wird in der generierten DB-Definition verwendet.
        /// </summary>
        public string Title { get; set; } = "_Alarms";

        /// <summary>
        /// Gibt an, ob der DB optimierten Zugriff verwendet.
        /// </summary>
        /// <remarks>
        /// Im TIA Portal ist für die Verwendung in HMI-Meldungen in der Regel <see langword="false"/> zu setzen.
        /// </remarks>
        public bool OptimizedAccess { get; set; } = false;

        /// <summary>
        /// Autor des generierten Datenbausteins.
        /// </summary>
        public string Author { get; set; } = "SIM-TIA-DeviceAlarmGenerator";

        /// <summary>
        /// Familienkennung für interne Zuordnung oder Bibliotheksversion.
        /// </summary>
        public string Family { get; set; } = "SIM";

        /// <summary>
        /// Versionsnummer der erzeugten Datenstruktur.
        /// </summary>
        public string Version { get; set; } = "0.1";

        /// <summary>
        /// Gibt an, ob der DB als nicht-retain deklariert wird.
        /// </summary>
        public bool NonRetain { get; set; } = true;

        // Standard-Alarme
        /// <summary>
        /// Liste der standardmäßig vorhandenen Alarme (z. B. Bibliotheksalarme).
        /// </summary>
        public List<AlarmBit> StandardAlarms { get; set; } = new();

        // Applikationsalarme
        /// <summary>
        /// Liste der applikationsspezifischen Alarme (z. B. maschinenspezifische Alarme).
        /// </summary>
        public List<AlarmBit> ApplicationAlarms { get; set; } = new();

        // Anzahl reservierter Quittierbits in Std_Ack
        /// <summary>
        /// Anzahl reservierter Quittierbits im Standard-Acknowledge-Bereich.
        /// </summary>
        public int StdAckReserveCount { get; set; } = 18;

        // DeviceSetting: Array[1..Rows, 0..2] of UInt
        /// <summary>
        /// Anzahl der Zeilen in der DeviceSetting-Matrix.
        /// </summary>
        public int DeviceSettingRows { get; set; } = 1;   // N

        /// <summary>
        /// Anzahl der Spalten in der DeviceSetting-Matrix.
        /// </summary>
        public int DeviceSettingCols { get; set; } = 3;   // 0..2

        /// <summary>
        /// Geräteeinstellungen als 2D-Array (UInt), z. B. zur Parametrierung der Geräte.
        /// </summary>  
        public uint[,] DeviceSetting { get; set; } = new uint[1, 3];

        // Max-Werte (werden im BEGIN geschrieben)
        /// <summary>
        /// Maximalwert der Geräteeinstellungen, wird im TIA-Quelltext im BEGIN-Bereich verwendet.
        /// </summary>
        public int DeviceSettingMax { get; set; } = 1;

        // Devices/Devices_Ack: Wortanzahl (Array [0..DevicesWordCount-1])
        /// <summary>
        /// Anzahl der Wörter im Devices-Array (z. B. Devices oder Devices_Ack).
        /// </summary>
        public int DevicesWordCount { get; set; } = 1;

        /// <summary>
        /// Maximalwert für Devices, entspricht der Wortanzahl.
        /// </summary>
        public int DevicesMax => DevicesWordCount;

        // ErrNoBuffer: Anzahl Einträge (Array [0..ErrNoBufferCount-1])
        /// <summary>
        /// Anzahl der Einträge im Fehlernummern-Puffer.
        /// </summary>
        public int ErrNoBufferCount { get; set; } = 10;

        /// <summary>
        /// Maximaler Index des Fehlernummern-Puffers.
        /// </summary>
        public int ErrNoBufferMax => ErrNoBufferCount - 1;

        // Std_Val (Werte, optional – nur Header/Struktur, keine Initialwerte)
        /// <summary>
        /// Standardwert-Blockstruktur, enthält Platzhalterwerte ohne Initialisierung.
        /// </summary>
        public StdValBlock StdVal { get; set; } = new();

        // NEU: Liste der Device-Typen aus StA-Cfg in Reihenfolge
        /// <summary>
        /// Liste aller im Projekt vorkommenden Gerätetypen in der Reihenfolge der StA-Cfg.
        /// </summary>
        public List<DeviceTypeInfo> DeviceTypes { get; set; } = new();

        /// <summary>
        /// Beinhaltet die Zuordnung von Gerätenamen (Devicetyp + Instanz) zu deren Label-Informationen.
        /// </summary>
        /// <remarks>
        /// Wird aus dem Excel-Sheet „Devices“ eingelesen und dient der korrekten Generierung von Alarmtexten
        /// pro Gerätetyp und Instanz.
        /// </remarks>
        public Dictionary<string, Dictionary<int, DeviceInstanceLabel>> DeviceInstanceLabels { get; set; }
            = new(StringComparer.OrdinalIgnoreCase);

        // ---- PREVIEW-PAYLOADS -----------------

        /// <summary>Rohdaten des Sheets „Devices“ als Matrix (nur für UI-Preview).</summary>
        public string[,]? DevicesMatrix { get; set; }

        /// <summary>Rohdaten des Sheets „StA-Cfg“ als Text-Matrix (nur für UI-Preview).</summary>
        public string[,]? StaCfgTextMatrix { get; set; }

        /// <summary>Rohdaten des Sheets „AppAlarm (_AA)“ als Matrix (nur für UI-Preview).</summary>
        public string[,]? AppAlarmMatrix { get; set; }
    }
}
