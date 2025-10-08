using SIM_TIA_DeviceAlarmgenerator.Model;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace SIM_TIA_DeviceAlarmgenerator.Services
{
    /// <summary>
    /// Baut aus <see cref="AlarmRow"/>-Sequenzen die finale 14-spaltige Matrix für den Excel-Export.
    /// </summary>
    /// <remarks>
    /// Die resultierende Matrix besitzt die Form <c>[AnzahlZeilen, 14]</c> und entspricht dem Ziel-Layout
    /// des HMI-/Excel-Exports (Spalten 0..13).
    /// </remarks>
    public interface IAlarmsMatrixBuilder
    {
        /// <summary>
        /// Wandelt erzeugte Alarmzeilen in eine 14-spaltige String-Matrix um.
        /// </summary>
        /// <param name="rows">Die Eingabesequenz von <see cref="AlarmRow"/>-Objekten.</param>
        /// <param name="ct">Optionaler <see cref="CancellationToken"/> zur vorzeitigen Beendigung.</param>
        /// <returns>
        /// Eine Task, die bei Abschluss eine 2D-String-Matrix der Größe <c>[AnzahlZeilen, 14]</c> liefert.
        /// </returns>
        /// <remarks>
        /// Typische Spaltenbelegung (0..13): RowNo, TagName, Message, (reserved), Class, Category,
        /// RefIndex, AckArea, AckIndex, Col9, Col10, Col11, Col12, Col13.
        /// </remarks>
        /// <exception cref="OperationCanceledException">
        /// Kann ausgelöst werden, wenn <paramref name="ct"/> den Abbruch signalisiert.
        /// </exception>
        Task<string[,]> BuildAsync(IEnumerable<AlarmRow> rows, CancellationToken ct = default);
    }
}
