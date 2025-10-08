using System;

namespace SIM_TIA_DeviceAlarmgenerator.Services
{
    /// <summary>
    /// Kontext, der der Nachrichten-Formatierung pro Zeile zur Verfügung gestellt wird.
    /// Enthält alle Indizes und Verweise, die zum Erzeugen des Meldungstextes benötigt werden.
    /// </summary>
    public readonly struct DeviceMessageContext
    {
        /// <summary>
        /// Erstellt einen neuen Kontext für die Nachrichtenzusammensetzung.
        /// </summary>
        /// <param name="globalRowNo1Based">Globale, 1-basierte Zeilennummer (entspricht der RowId im Textprefix).</param>
        /// <param name="deviceIndex1Based">Geräteindex 1-basiert (Instanznummer innerhalb des Gerätetyps).</param>
        /// <param name="alarmIndex0Based">Alarmindex 0-basiert (Position des Bits innerhalb des Typschemas).</param>
        /// <param name="devicesMatrix">Matrix des Sheets „Devices“ im Format <c>[Zeile, Spalte]</c>.</param>
        /// <param name="dataStartRow1Based">1-basiger Startindex des Datenbereichs für den gefundenen Gerätetyp-Block.</param>
        /// <param name="blockSpan">Gesamtausdehnung des gefundenen Blocks (inkl. Kopf/Trennerzeilen).</param>
        /// <param name="elementsFromTab">Anzahl nutzbarer Datenelemente, die aus der „Devices“-Tabelle stammen.</param>
        public DeviceMessageContext(
            int globalRowNo1Based,
            int deviceIndex1Based,
            int alarmIndex0Based,
            string[,] devicesMatrix,
            int dataStartRow1Based,
            int blockSpan,
            int elementsFromTab)
        {
            GlobalRowNo1Based = globalRowNo1Based;
            DeviceIndex1Based = deviceIndex1Based;
            AlarmIndex0Based = alarmIndex0Based;
            DevicesMatrix = devicesMatrix;
            DataStartRow1Based = dataStartRow1Based;
            BlockSpan = blockSpan;
            ElementsFromTab = elementsFromTab;
        }

        /// <summary>
        /// Globale, 1-basierte Zeilennummer (wird typischerweise dem Meldungstext vorangestellt).
        /// </summary>
        public int GlobalRowNo1Based { get; }

        /// <summary>
        /// 1-basierter Geräteindex (Instanz innerhalb des Gerätetyps).
        /// </summary>
        public int DeviceIndex1Based { get; }

        /// <summary>
        /// 0-basierter Alarmindex innerhalb des Geräteschemas/Fehlervorlage.
        /// </summary>
        public int AlarmIndex0Based { get; }

        /// <summary>
        /// Referenz auf die „Devices“-Matrix (<c>[Zeile, Spalte]</c>) zur Label-/Namensauflösung.
        /// </summary>
        public string[,] DevicesMatrix { get; }

        /// <summary>
        /// 1-basiger Startindex des Datenbereichs (erste Datenzeile nach dem Header des Gerätetyp-Blocks).
        /// </summary>
        public int DataStartRow1Based { get; }

        /// <summary>
        /// Gesamtspanne des gefundenen Blocks (inklusive Kopf-/Trennerzeilen).
        /// </summary>
        public int BlockSpan { get; }

        /// <summary>
        /// Anzahl der aus der Tabelle tatsächlich nutzbaren Elemente (z. B. vorhandene Instanzen).
        /// </summary>
        public int ElementsFromTab { get; }
    }

    /// <summary>
    /// Strategie/Regel für einen Device-Typ (Name, Fehlervorlage, Alarme/Device, Nachrichtenaufbau).
    /// </summary>
    public interface IDeviceAlarmRule
    {
        /// <summary>
        /// Geräte-Typname genau wie in Spalte 0 im „Devices“-Sheet (z. B. „Omode“).
        /// </summary>
        string DeviceTypeName { get; }

        /// <summary>
        /// Anzahl Alarme pro Device dieses Typs.
        /// </summary>
        int AlarmsPerDevice { get; }

        /// <summary>
        /// Liste der (Text, Klasse)-Paare je Alarmindex (Index = 0-basiert).
        /// Die Klasse ist in der Regel „Errors“ oder „Warnings“.
        /// </summary>
        (string Text, string Class)[] ErrorsTemplate { get; }

        /// <summary>
        /// Liefert den Meldungstext für eine konkrete Zeile (Spalte 2 im Ziel-Layout).
        /// </summary>
        /// <param name="ctx">Kontext mit Matrixzugriff und Indizes.</param>
        /// <returns>Der formatierte Meldungstext.</returns>
        string ComposeMessage(DeviceMessageContext ctx);

        /// <summary>
        /// Optional: Liefert den <c>TagName</c> (Spalte 1) für eine konkrete Zeile.
        /// Standardformat: <c>{DeviceTypeName}_{AlarmIndex0Based}_{DeviceIndex1Based:000}</c>.
        /// </summary>
        /// <param name="alarmIndex0Based">Alarmindex (0-basiert).</param>
        /// <param name="deviceIndex1Based">Geräteindex (1-basiert, dreistellig im TagName).</param>
        /// <returns>TagName der Zeile.</returns>
        string ComposeTagName(int alarmIndex0Based, int deviceIndex1Based) =>
            $"{DeviceTypeName}_{alarmIndex0Based}_{(deviceIndex1Based < 10 ? $"00{deviceIndex1Based}" : (deviceIndex1Based < 100 ? $"0{deviceIndex1Based}" : deviceIndex1Based.ToString()))}";
    }
}
