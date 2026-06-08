using CommunityToolkit.Mvvm.ComponentModel;
using SIM_TIA_DeviceAlarmgenerator.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;

namespace SIM_TIA_DeviceAlarmgenerator.ViewModel
{
    public partial class AlarmDbPreviewViewModel : ObservableObject
    {
        [ObservableProperty]
        private AlarmDbModel _model;

        // DataTables für 2D-Matrizen (bequem für DataGrid)
        public DataTable DeviceSettingTable { get; }
        public DataTable DevicesMatrixTable { get; }
        public DataTable StaCfgTextMatrixTable { get; }
        public DataTable AppAlarmMatrixTable { get; }

        // Tree für Instance Labels
        public ObservableCollection<DeviceGroupNode> InstanceLabelTree { get; } = new();

        public AlarmDbPreviewViewModel() 
        { 
            Model = new AlarmDbModel();
            DeviceSettingTable = new DataTable();
            DevicesMatrixTable = new DataTable();
            StaCfgTextMatrixTable = new DataTable();
            AppAlarmMatrixTable = new DataTable();
        }

        public AlarmDbPreviewViewModel(AlarmDbModel model)
        {
            Model = model ?? throw new ArgumentNullException(nameof(model));

            DeviceSettingTable = BuildTableFromUInt2D(Model.DeviceSetting);
            DevicesMatrixTable = BuildTableFromString2D(Model.DevicesMatrix);
            StaCfgTextMatrixTable = BuildTableFromString2D(Model.StaCfgTextMatrix);
            AppAlarmMatrixTable = BuildTableFromString2D(Model.AppAlarmMatrix);

            BuildInstanceLabelTree(Model.DeviceInstanceLabels);
        }

        private static DataTable BuildTableFromUInt2D(uint[,]? data)
        {
            var table = new DataTable("DeviceSetting");
            if (data == null) return table;

            int rows = data.GetLength(0);
            int cols = data.GetLength(1);
            for (int c = 0; c < cols; c++)
                table.Columns.Add($"C{c}", typeof(uint));

            for (int r = 0; r < rows; r++)
            {
                var row = table.NewRow();
                for (int c = 0; c < cols; c++)
                    row[c] = data[r, c];
                table.Rows.Add(row);
            }
            return table;
        }

        private static DataTable BuildTableFromString2D(string[,]? data)
        {
            var table = new DataTable("Matrix");
            if (data == null) return table;

            int rows = data.GetLength(0);
            int cols = data.GetLength(1);
            for (int c = 0; c < cols; c++)
                table.Columns.Add($"C{c}", typeof(string));

            for (int r = 0; r < rows; r++)
            {
                var row = table.NewRow();
                for (int c = 0; c < cols; c++)
                    row[c] = data[r, c];
                table.Rows.Add(row);
            }
            return table;
        }

        private void BuildInstanceLabelTree(
            Dictionary<string, Dictionary<int, DeviceInstanceLabel>> dict)
        {
            InstanceLabelTree.Clear();
            if (dict == null) return;

            foreach (var (deviceType, instances) in dict)
            {
                var group = new DeviceGroupNode(deviceType);
                foreach (var (instance, label) in instances)
                    group.Instances.Add(new DeviceInstanceNode(instance, label));
                InstanceLabelTree.Add(group);
            }
        }
    }

    // --- Tree-Knoten für Instance Labels ---
    public sealed class DeviceGroupNode
    {
        public string DeviceType { get; }
        public ObservableCollection<DeviceInstanceNode> Instances { get; } = new();
        public DeviceGroupNode(string deviceType) => DeviceType = deviceType;
    }

    public sealed class DeviceInstanceNode
    {
        public int Instance { get; }
        public DeviceInstanceLabel Label { get; }
        public DeviceInstanceNode(int instance, DeviceInstanceLabel label)
        {
            Instance = instance;
            Label = label;
        }
    }
}
