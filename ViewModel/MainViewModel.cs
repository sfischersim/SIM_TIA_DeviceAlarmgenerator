using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using log4net;
using SIM_TIA_DeviceAlarmgenerator.Model;
using SIM_TIA_DeviceAlarmgenerator.Services;
using SIM_TIA_DeviceAlarmgenerator.Services.DeviceRules;
using System;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace SIM_TIA_DeviceAlarmgenerator.ViewModel
{
    /// <summary>
    /// Haupt-ViewModel der Anwendung. Orchestriert das Einlesen der Excel-Konfiguration,
    /// den Aufbau der Alarmmodelle und die Exporte (Siemens <c>.db</c> und HMI-Alarme als Excel).
    /// </summary>
    /// <remarks>
    /// Verwendet CommunityToolkit.Mvvm für Property-/Command-Generierung sowie WPF-Dispatcher
    /// für UI-sichere Status-Updates.  
    /// Die in diesem ViewModel erzeugten Statusmeldungen werden in <see cref="LogEntries"/> protokolliert.
    /// </remarks>
    internal partial class MainViewModel : ObservableObject
    {
        private ILog log = App.GetLogger();

        private readonly FilePickerService? _filePicker = new FilePickerService();
        private readonly IAlarmsMatrixBuilder _matrixBuilder = new AlarmsMatrixBuilder();
        /// <summary>
        /// Chronologisch absteigende Sammlung an Status-/Logeinträgen für die UI.
        /// </summary>
        public ObservableCollection<string> LogEntries { get; } = new();
        private readonly Dispatcher _dispatcher = Dispatcher.CurrentDispatcher;
        private readonly ExcelAlarmConfigReader _excelReader = new();
        private readonly HmiAlarmsExporter _hmiExporter = new();
        private readonly IAlarmDbExporter _dbExporter = new AlarmDbExporter();

        // -----------------------------
        // Properties
        // -----------------------------

        /// <summary>Vollständiger Pfad zur Projekt-Excel (xlsx/xlsm).</summary>
        [ObservableProperty]
        private string? filePath;

        /// <summary>Matrix des Blatts <c>Devices</c>.</summary>
        [ObservableProperty]
        private string[,]? devicesMatrix;

        /// <summary>Matrix des Blatts <c>StA-Cfg</c> [30,2].</summary>
        [ObservableProperty]
        private int[,]? staCfgMatrix;

        /// <summary>Text-Matrix des Blatts <c>StA-Cfg</c> (für Mengenermittlung je Device-Typ).</summary>
        [ObservableProperty]
        private string[,]? staCfgTextMatrix;

        /// <summary>Matrix des Blatts <c>AppAlarm (_AA)</c>.</summary>
        [ObservableProperty]
        private string[,]? appAlarmMatrix;

        /// <summary>Statusmeldung für die UI.</summary>
        [ObservableProperty]
        private string status = "Bereit";

        /// <summary>Vorschau als DataView: Devices.</summary>
        [ObservableProperty]
        private DataView? devicesPreview;

        /// <summary>Vorschau als DataView: StA-Cfg.</summary>
        [ObservableProperty]
        private DataView? numbersPreview;

        /// <summary>Vorschau als DataView: AppAlarm.</summary>
        [ObservableProperty]
        private DataView? appAlarmPreview;

        /// <summary>
        /// Zuletzt exportierter Alarms-DB (voller Pfad).
        /// </summary>
        [ObservableProperty]
        private string? lastGeneratedDbPath;

        /// <summary>
        /// Fertige Alarm-Tabelle (jede Zeile = ein Datensatz).  
        /// Spaltenlayout: 14 Spalten entsprechend HMI-/Export-Layout (0..13).
        /// </summary>
        [ObservableProperty]
        private string[,]? alarmsMatrix;

        /// <summary>
        /// Zuletzt exportierte Excel-Datei (voller Pfad).
        /// </summary>
        [ObservableProperty]
        private string? lastExportPath;

        /// <summary>
        /// Steuert ProgressBar-Visibility.
        /// </summary>
        [ObservableProperty]
        private bool isBusy;

        /// <summary>
        /// Das aktuell geladene Alarm-Datenmodell, aufgebaut aus der Excel-Konfiguration.
        /// </summary>
        [ObservableProperty]
        private AlarmDbModel alarmDb = new AlarmDbModel();

        [ObservableProperty]
        public AlarmDbPreviewViewModel previewVm = new AlarmDbPreviewViewModel();

        /// <summary>
        /// Öffnet einen Dateiauswahldialog und liest die gewählte Projekt-Excel ein.
        /// </summary>
        /// <param name="ct">Abbruch-Token für die asynchrone Ausführung.</param>
        [RelayCommand]
        private async Task PickAndReadAsync(CancellationToken ct)
        {
            SetStatus("Dateiauswahl gestartet");
            IsBusy = true;
            try
            {
                FilePath = _filePicker?.PickExcelFile();
                if (!string.IsNullOrWhiteSpace(FilePath))
                {
                    ReadData();
                    SetStatus($"Datei {FilePath} geöffnet");
                }
                else
                {
                    SetStatus("Kein Datei-Pfad ausgewählt.");
                }

                IsBusy = false;
            }
            catch (Exception ex)
            {
                SetStatus($"Fehler beim Öffnen: {ex.Message}");
                IsBusy = false;
            }
        }

        /// <summary>
        /// Exportiert die aktuelle Alarm-DB als Siemens-<c>.db</c>-Quelldatei.
        /// </summary>
        [RelayCommand]
        private void ExportAlarmDb()
        {
            var sfd = new Microsoft.Win32.SaveFileDialog
            {
                Title = "alarm.db exportieren",
                Filter = "Siemens DB (*.db)|*.db|Alle Dateien (*.*)|*.*",
                FileName = $"{AlarmDb?.Name ?? "_Alarms"}.db",
                AddExtension = true,
                OverwritePrompt = true
            };
            if (sfd.ShowDialog() != true) return;

            try
            {
                if (AlarmDb == null)
                {
                    SetStatus("Kein Datenmodell vorhanden (bitte Excel einlesen).");
                    return;
                }
                SetStatus("Alarm-DB-Export gestartet");
                _dbExporter.ExportAsync(AlarmDb, sfd.FileName).GetAwaiter().GetResult();
                LastGeneratedDbPath = sfd.FileName;
                SetStatus($"Alarm-DB exportiert: {sfd.FileName}");
            }
            catch (Exception ex)
            {
                SetStatus($"Fehler beim Alarm-DB-Export: {ex.Message}");
            }
        }

        /// <summary>
        /// Exportiert die HMI-Alarme als Excel (Sheet „DiscreteAlarms“).
        /// </summary>
        [RelayCommand]
        private void ExportHmiAlarms()
        {
            if (AlarmDb == null)
            {
                SetStatus("Keine Daten für HMI-Export (bitte Excel einlesen).");
                return;
            }

            var sfd = new Microsoft.Win32.SaveFileDialog
            {
                Title = "HMI-Alarme (Excel) exportieren",
                Filter = "Excel-Datei (*.xlsx)|*.xlsx",
                FileName = $"HMI_Alarms_{DateTime.Now:yyyyMMdd_HHmm}.xlsx",
                AddExtension = true,
                OverwritePrompt = true
            };
            if (sfd.ShowDialog() != true) return;

            try
            {
                SetStatus("HMI-Alarm-Export gestartet");

                // Hinweis: DeviceTypes werden bereits im AlarmDbModel aus der Excel gelesen.
                if (AlarmDb.DeviceTypes == null || AlarmDb.DeviceTypes.Count == 0)
                {
                    SetStatus("WARNUNG: Keine DeviceTypes im Model – Export würde leer sein.");
                    return;
                }

                _hmiExporter.Export(sfd.FileName, AlarmDb);
                LastExportPath = sfd.FileName;
                SetStatus($"HMI-Alarme exportiert: {sfd.FileName}");
            }
            catch (Exception ex)
            {
                SetStatus($"Fehler beim HMI-Export: {ex.Message}");
            }
        }


        /// <summary>
        /// Liest Mengen aus „StA-Cfg“ und erzeugt Alarme gemäß der erkannten <see cref="IDeviceAlarmRule"/>-Regeln.
        /// </summary>
        /// <param name="ct">Abbruch-Token für die asynchrone Ausführung.</param>
        [RelayCommand]
        private async Task BuildAlarmsFromStaCfgAsync(CancellationToken ct)
        {
            SetStatus("Baue Alarme zusammen.");
            IsBusy = true;

            if (DevicesMatrix is null || StaCfgMatrix is null)
            {
                SetStatus("Bitte zuerst Devices und StA-Cfg einlesen.");
                IsBusy = false;
                return;
            }

            var qtyProvider = new StaCfgQuantityProvider(StaCfgTextMatrix);

            var rules = DeviceAlarmRuleFactory.GetAllRules();

            var rows = new System.Collections.Generic.List<AlarmRow>();

            foreach (var rule in rules)
            {
                // Mengen automatisch aus StA-Cfg ziehen (inkl. Normalisierung)
                if (!qtyProvider.TryGetQuantity(rule.DeviceTypeName, out var qty) || qty <= 0)
                {
                    // Fallback: ohne Unterstrich probieren
                    var altKey = rule.DeviceTypeName.TrimEnd('_', ' ');
                    qtyProvider.TryGetQuantity(altKey, out qty);
                }

                if (qty > 0)
                {
                    var builder = new GenericDeviceAlarmBuilder(rule);
                    rows.AddRange(builder.Build(DevicesMatrix, qty));
                }
            }

            AlarmsMatrix = await _matrixBuilder.BuildAsync(rows, ct);
            SetStatus($"Erzeugte Gesamtzeilen: {AlarmsMatrix.GetLength(0)}");

            IsBusy = false;
        }

        /// <summary>
        /// Liest die ausgewählte Excel-Datei ein und baut das <see cref="AlarmDb"/>-Modell.
        /// </summary>
        private void ReadData()
        {
            try
            {
                SetStatus("Excel lesen gestartet");
                if (string.IsNullOrWhiteSpace(FilePath))
                {
                    SetStatus("Kein Datei-Pfad gesetzt.");
                    return;
                }

                // 1) Vollmodell aufbauen
                var model = _excelReader.BuildFromExcel(FilePath);
                AlarmDb = model;

                // 2) Previews holen
                DevicesPreview = ToDataView(model.DevicesMatrix);
                NumbersPreview = ToDataView(model.StaCfgTextMatrix);
                AppAlarmPreview = ToDataView(model.AppAlarmMatrix);

                // 3) kurze Statistiken
                if (DevicesPreview != null) SetStatus($"Devices: {DevicesPreview.Count} Zeilen");
                if (NumbersPreview != null) SetStatus($"StA-Cfg: {NumbersPreview.Count} Zeilen");
                if (AppAlarmPreview != null) SetStatus($"AppAlarm (_AA): {AppAlarmPreview.Count} Zeilen");

                // 4) Zusammenfassung aus deinem Model (existiert bei dir bereits)
                SetStatus($"Excel gelesen: DeviceTypeCount={model.DeviceSettingRows}, DevicesMax={model.DevicesMax}, ErrNoBufferMax={model.ErrNoBufferMax}");

                // 5) Preview erzeugen
                PreviewVm = new AlarmDbPreviewViewModel(model);
            }
            catch (Exception ex)
            {
                SetStatus($"Fehler beim Excel-Einlesen: {ex.Message}");
            }
        }

        #region Log & Status

        /// <summary>
        /// Setzt eine Statusmeldung und fügt sie dem Log hinzu (thread-sicher über den Dispatcher).
        /// </summary>
        /// <param name="message">Die auszugebende Statusnachricht.</param>
        private void SetStatus(string message)
        {
            var stamp = DateTime.Now.ToString("HH:mm:ss");
            var line = $"[{stamp}] {message}";

            // auf UI-Thread bringen, falls Aufruf aus Task/Thread
            if (!_dispatcher.CheckAccess())
            {
                _dispatcher.Invoke(() => ApplyStatus(line));
            }
            else
            {
                ApplyStatus(line);
            }
        }

        /// <summary>
        /// Wendet die Statusmeldung auf UI-gebundene Properties an und pflegt das Log.
        /// </summary>
        /// <param name="line">Vollständige Logzeile (inklusive Zeitstempel).</param>
        private void ApplyStatus(string line)
        {
            Status = line;

            // Neuesten Eintrag oben hinzufügen
            LogEntries.Insert(0, line);

            // Log begrenzen (älteste am Ende löschen)
            const int max = 500;
            if (LogEntries.Count > max)
                LogEntries.RemoveAt(LogEntries.Count - 1);
        }

        /// <summary>
        /// Löscht die aktuelle Logliste und setzt den Status entsprechend.
        /// </summary>
        [RelayCommand]
        private void ClearLog()
        {
            LogEntries.Clear();
            SetStatus("Log geleert");
        }

        #endregion

        #region Helper
        private static DataView? ToDataView(string[,]? matrix)
        {
            if (matrix is null) return null;
            var rows = matrix.GetLength(0);
            var cols = matrix.GetLength(1);

            var table = new DataTable();
            for (int c = 0; c < cols; c++)
                table.Columns.Add($"C{c + 1}", typeof(string));

            for (int r = 0; r < rows; r++)
            {
                var dr = table.NewRow();
                for (int c = 0; c < cols; c++)
                    dr[c] = matrix[r, c];
                table.Rows.Add(dr);
            }
            return table.DefaultView;
        }

        private static DataView? ToDataView(DataTable? table)
        {
            return table?.DefaultView;
        }

        private static DataView? ToDataView<T>(IEnumerable<T>? rows)
        {
            if (rows is null) return null;

            var table = new DataTable(typeof(T).Name);

            // Spalten aus öffentlichen Properties ableiten
            var props = typeof(T).GetProperties(System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public);
            foreach (var p in props)
            {
                var colType = Nullable.GetUnderlyingType(p.PropertyType) ?? p.PropertyType;
                table.Columns.Add(p.Name, colType);
            }

            // Daten füllen
            foreach (var item in rows)
            {
                var dr = table.NewRow();
                foreach (var p in props)
                {
                    var value = p.GetValue(item, null);
                    dr[p.Name] = value ?? DBNull.Value;
                }
                table.Rows.Add(dr);
            }

            return table.DefaultView;
        }

        #endregion
    }
}
