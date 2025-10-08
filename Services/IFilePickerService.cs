using Microsoft.Win32;

namespace SIM_TIA_DeviceAlarmgenerator.Services
{
    /// <summary>
    /// Abstraktion für das Auswählen einer Excel-Datei (.xlsx oder .xlsm).  
    /// Diese Schnittstelle ermöglicht es, das ViewModel vom UI zu entkoppeln und somit testbar zu halten.
    /// </summary>
    public interface IFilePickerService
    {
        /// <summary>
        /// Öffnet einen Dateiauswahldialog und gibt den ausgewählten Pfad zurück.  
        /// Gibt <see langword="null"/> zurück, wenn der Benutzer den Dialog abbricht.
        /// </summary>
        /// <returns>Vollständiger Dateipfad der ausgewählten Excel-Datei oder <see langword="null"/>.</returns>
        string? PickExcelFile();
    }
}
