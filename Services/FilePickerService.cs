using Microsoft.Win32;
using System;

namespace SIM_TIA_DeviceAlarmgenerator.Services
{
    /// <summary>
    /// WPF-Implementierung des <see cref="IFilePickerService"/> auf Basis von <see cref="OpenFileDialog"/>.
    /// </summary>
    /// <remarks>
    /// Unterstützt die Auswahl von Excel-Dateien (<c>.xlsx</c>, <c>.xlsm</c>).  
    /// Aufruf sollte vom UI-Thread erfolgen (WPF-Dialog).
    /// </remarks>
    public sealed class FilePickerService : IFilePickerService
    {
        /// <summary>
        /// Öffnet einen Dateiauswahldialog für Excel-Dateien und gibt den ausgewählten Pfad zurück.
        /// </summary>
        /// <returns>
        /// Vollständiger Dateipfad der ausgewählten Datei oder <see langword="null"/>,
        /// wenn der Dialog abgebrochen wurde.
        /// </returns>
        /// <remarks>
        /// Verwendet den Filter „*.xlsx;*.xlsm“ und setzt <c>DefaultExt</c> auf „.xlsm“.  
        /// Mehrfachauswahl ist deaktiviert.
        /// </remarks>
        public string? PickExcelFile()
        {
            var dlg = new OpenFileDialog
            {
                Title = "Excel-Datei auswählen",
                Filter = "Excel Dateien (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Alle Dateien (*.*)|*.*",
                DefaultExt = ".xlsm",
                CheckFileExists = true,
                CheckPathExists = true,
                Multiselect = false
            };

            return dlg.ShowDialog() == true ? dlg.FileName : null;
        }
    }
}
