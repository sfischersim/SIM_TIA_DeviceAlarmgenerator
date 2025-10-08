using System.Collections.Generic;

namespace SIM_TIA_DeviceAlarmgenerator.Services
{
    /// <summary>
    /// Bietet Zugriff auf Geräte- und Alarm-Mengen aus einer Konfigurationstabelle (z. B. „StA-Cfg“).
    /// </summary>
    /// <remarks>
    /// Diese Schnittstelle stellt lesenden Zugriff auf alle Gerätetypen, deren konfigurierte Geräteanzahl (Quantity)
    /// und die jeweilige Alarmanzahl pro Gerät bereit.  
    /// Sie dient insbesondere als Datenquelle für den Aufbau von Alarmdatenbanken oder Exportstrukturen.
    /// </remarks>
    public interface IDeviceQuantityProvider
    {
        /// <summary>
        /// Liefert die konfigurierte Geräteanzahl (Quantity) für den angegebenen Device-Typ.
        /// </summary>
        /// <param name="deviceTypeName">Gerätename wie im „Devices“-Sheet (z. B. „Omode“, „2DS_“, „SQ“).</param>
        /// <param name="quantity">Ausgegebene Geräteanzahl oder 0, wenn nicht gefunden oder nicht parsbar.</param>
        /// <returns><see langword="true"/>, wenn ein Eintrag gefunden wurde; andernfalls <see langword="false"/>.</returns>
        bool TryGetQuantity(string deviceTypeName, out int quantity);

        /// <summary>
        /// Liefert die konfigurierte Alarmanzahl pro Gerät (falls vorhanden) für den angegebenen Gerätetyp.
        /// </summary>
        /// <param name="deviceTypeName">Gerätename wie im „Devices“-Sheet.</param>
        /// <param name="alarmsPerDevice">Ausgegebene Alarmanzahl oder 0, wenn kein Wert vorhanden ist.</param>
        /// <returns><see langword="true"/>, wenn ein Eintrag gefunden wurde; andernfalls <see langword="false"/>.</returns>
        bool TryGetAlarmsPerDevice(string deviceTypeName, out int alarmsPerDevice);

        /// <summary>
        /// Liefert ein schreibgeschütztes Mapping aller bekannten Gerätetypen
        /// mit ihrer jeweiligen Geräteanzahl und Alarmanzahl pro Gerät.
        /// </summary>
        /// <remarks>
        /// Schlüssel: Geräte-Typname  
        /// Wert: Tupel (<c>Quantity</c>, <c>AlarmsPerDevice</c>)
        /// </remarks>
        IReadOnlyDictionary<string, (int Quantity, int AlarmsPerDevice)> All { get; }
    }
}
