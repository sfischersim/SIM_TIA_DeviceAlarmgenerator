using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SIM_TIA_DeviceAlarmgenerator.Services
{
    /// <summary>
    /// Liest Geräte- und Alarmmengen aus der „StA-Cfg“-Matrix.
    /// Erwartet: Spalte 1 = Device-Typ, Spalte 2 = Device Quantity, Spalte 3 = Alarm Quantity (0-basiert).
    /// </summary>
    public sealed class StaCfgQuantityProvider : IDeviceQuantityProvider
    {
        private readonly Dictionary<string, (int Quantity, int AlarmsPerDevice)> _map;

        /// <summary>
        /// Erstellt einen Provider auf Basis einer gelesenen „StA-Cfg“-Matrix.
        /// </summary>
        /// <param name="staCfgMatrix">
        /// Matrix <c>[Zeile,Spalte]</c> des „StA-Cfg“-Sheets.
        /// Erwartet Spalten: 1 = Typ, 2 = Geräteanzahl, 3 = Alarme/Device (0-basiert).
        /// </param>
        /// <param name="deviceTypeCol">Index der Typ-Spalte (Standard: 1).</param>
        /// <param name="quantityCol">Index der Geräteanzahl-Spalte (Standard: 2).</param>
        /// <param name="alarmsCol">Index der Alarmanzahl-Spalte (Standard: 3).</param>
        /// <exception cref="ArgumentNullException"><paramref name="staCfgMatrix"/> ist <see langword="null"/>.</exception>
        public StaCfgQuantityProvider(
            string[,] staCfgMatrix,
            int deviceTypeCol = 1,
            int quantityCol = 2,
            int alarmsCol = 3)
        {
            if (staCfgMatrix is null) throw new ArgumentNullException(nameof(staCfgMatrix));

            _map = new Dictionary<string, (int, int)>(StringComparer.OrdinalIgnoreCase);

            int rows = staCfgMatrix.GetLength(0);
            int cols = staCfgMatrix.GetLength(1);

            for (int r = 0; r < rows; r++)
            {
                string typeRaw = Safe(staCfgMatrix, r, deviceTypeCol);
                if (string.IsNullOrWhiteSpace(typeRaw)) continue;

                string key = Normalize(typeRaw);

                int qty = ParseIntSafe(Safe(staCfgMatrix, r, quantityCol));
                int alm = ParseIntSafe(Safe(staCfgMatrix, r, alarmsCol));

                // Wenn mehrfach vorhanden: letzter Eintrag gewinnt (typisch unkritisch)
                _map[key] = (qty, alm);
            }
        }

        /// <inheritdoc />
        public bool TryGetQuantity(string deviceTypeName, out int quantity)
        {
            quantity = 0;
            if (string.IsNullOrWhiteSpace(deviceTypeName)) return false;

            if (_map.TryGetValue(Normalize(deviceTypeName), out var tuple))
            {
                quantity = tuple.Quantity;
                return true;
            }
            return false;
        }

        /// <inheritdoc />
        public bool TryGetAlarmsPerDevice(string deviceTypeName, out int alarmsPerDevice)
        {
            alarmsPerDevice = 0;
            if (string.IsNullOrWhiteSpace(deviceTypeName)) return false;

            if (_map.TryGetValue(Normalize(deviceTypeName), out var tuple))
            {
                alarmsPerDevice = tuple.AlarmsPerDevice;
                return true;
            }
            return false;
        }

        /// <inheritdoc />
        public IReadOnlyDictionary<string, (int Quantity, int AlarmsPerDevice)> All => _map;

        /// <summary>
        /// Liefert einen Zellenwert index-sicher (leer, wenn außerhalb oder <see langword="null"/>).
        /// </summary>
        private static string Safe(string[,] m, int r, int c)
        {
            if (r < 0 || c < 0) return string.Empty;
            if (r >= m.GetLength(0) || c >= m.GetLength(1)) return string.Empty;
            return m[r, c] ?? string.Empty;
        }

        /// <summary>
        /// Parsen von Ganzzahlen, tolerant gegenüber Leerstrings und Kultur.
        /// </summary>
        private static int ParseIntSafe(string s)
            => int.TryParse(s?.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : 0;

        /// <summary>
        /// Normalisiert Typ-Namen, um Schreibvarianten zusammenzuführen (z. B. „2DS“ ↔ „2DS_“).
        /// </summary>
        private static string Normalize(string s)
        {
            s = s.Trim();
            // Entferne trailing Unterstriche oder Leerzeichen
            while (s.EndsWith("_", StringComparison.Ordinal) || s.EndsWith(" ", StringComparison.Ordinal))
                s = s.Substring(0, s.Length - 1);

            return s;
        }
    }
}
