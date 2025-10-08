using SIM_TIA_DeviceAlarmgenerator.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SIM_TIA_DeviceAlarmgenerator.Services
{
    /// <inheritdoc />
    /// <summary>
    /// Baut aus einer Liste von <see cref="AlarmRow"/> eine 14-spaltige Zeichenfolgenmatrix
    /// im Excel-Ziel­layout (Spalten 0..13) mit fortlaufender globaler Nummerierung.
    /// </summary>
    /// <remarks>
    /// Die Methode setzt den Meldungstext (Spalte 2) so um, dass eine führende Nummer vor dem Doppelpunkt
    /// durch die globale lfd. Nummer ersetzt wird. Existiert kein Doppelpunkt, wird „{Nr}: “ vorangestellt.
    /// </remarks>
    public sealed class AlarmsMatrixBuilder : IAlarmsMatrixBuilder
    {
        /// <summary>
        /// Erzeugt die 14-spaltige Matrix aus den übergebenen Alarmzeilen.
        /// </summary>
        /// <param name="rows">Quellzeilen. Bei <see langword="null"/> wird eine leere Liste verwendet.</param>
        /// <param name="ct">Abbruch-Token zur vorzeitigen Beendigung der Verarbeitung.</param>
        /// <returns>Eine Matrix der Größe <c>[Anzahl Zeilen, 14]</c> im Excel-Ziel­layout.</returns>
        /// <remarks>
        /// Spalte 0 enthält die globale lfd. Nummer; Spalte 2 den ggf. umgeschriebenen Meldungstext.
        /// Alle anderen Spalten werden 1:1 aus <see cref="AlarmRow"/> übernommen.
        /// </remarks>
        /// <exception cref="OperationCanceledException">
        /// Wird ausgelöst, wenn <paramref name="ct"/> den Abbruch signalisiert.
        /// </exception>
        public Task<string[,]> BuildAsync(IEnumerable<AlarmRow> rows, CancellationToken ct = default)
        {
            var list = rows?.ToList() ?? new List<AlarmRow>();
            const int cols = 14;
            var matrix = new string[list.Count, cols];

            int globalNr = 1;

            for (int r = 0; r < list.Count; r++)
            {
                ct.ThrowIfCancellationRequested();
                var a = list[r];

                matrix[r, 0] = globalNr.ToString();

                // Message ggf. neu zusammensetzen: führende Zahl durch globale Nummer ersetzen
                string msg = a.Message ?? string.Empty;
                int colonIdx = msg.IndexOf(':');
                if (colonIdx > 0)
                {
                    msg = $"{globalNr}{msg.Substring(colonIdx)}";
                }
                else
                {
                    // falls keine führende Zahl mit Doppelpunkt existiert
                    msg = $"{globalNr}: {msg}";
                }
                matrix[r, 2] = msg;

                // Rest bleibt gleich
                matrix[r, 1] = a.TagName;
                matrix[r, 3] = a.Reserved3;
                matrix[r, 4] = a.Class;
                matrix[r, 5] = a.Category;
                matrix[r, 6] = a.RefIndex;
                matrix[r, 7] = a.AckArea;
                matrix[r, 8] = a.AckIndex;
                matrix[r, 9] = a.Col9;
                matrix[r, 10] = a.Col10;
                matrix[r, 11] = a.Col11;
                matrix[r, 12] = a.Col12;
                matrix[r, 13] = a.Col13;

                globalNr++;
            }

            return Task.FromResult(matrix);
        }
    }
}
