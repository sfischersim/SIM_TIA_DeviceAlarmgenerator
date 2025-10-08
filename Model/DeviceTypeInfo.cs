using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SIM_TIA_DeviceAlarmgenerator.Model
{
    /// <summary>
    /// Enthält die Typinformationen eines Geräts, das im Alarmdatenbaustein berücksichtigt wird.
    /// </summary>
    /// <remarks>
    /// Diese Klasse beschreibt den Gerätenamen, die Anzahl der Instanzen und die Anzahl der Alarme je Gerätetyp.
    /// Sie dient als Basis für die Generierung der Gerätekonfiguration und der zugehörigen Alarmstrukturen.
    /// </remarks>
    public sealed class DeviceTypeInfo
    {
        /// <summary>
        /// Name des Gerätetyps, z. B. „APosi“, „BC“ oder „CVX“.
        /// </summary>
        /// <remarks>
        /// Entspricht dem internen Bezeichner des Device-Typs im Standard oder der StA-Cfg.
        /// </remarks>
        public string DevName { get; set; } = "";

        /// <summary>
        /// Anzahl der im Projekt vorhandenen Geräte dieses Typs.
        /// </summary>
        /// <remarks>
        /// Wird verwendet, um die Größe der Device-Arrays und die Anzahl der Alarmbits zu bestimmen.
        /// </remarks>
        public int DevQty { get; set; }

        /// <summary>
        /// Anzahl der Alarme pro Gerät dieses Typs.
        /// </summary>
        /// <remarks>
        /// Bestimmt, wie viele Alarmbits pro Instanz im Alarmdatenbaustein reserviert werden.
        /// </remarks>
        public int AlmQty { get; set; }
    }
}
