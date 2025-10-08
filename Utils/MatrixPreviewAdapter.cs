using System;
using System.Data;

namespace SIM_TIA_DeviceAlarmgenerator.Utils
{
    /// <summary>
    /// Wandelt 2D-Matrizen (<see cref="string[,]"/> / <see cref="int[,]"/>) 
    /// in <see cref="DataView"/> um, ohne Excel-Header oder Zeilennummern.
    /// </summary>
    public static class MatrixPreviewAdapter
    {
        /// <summary>
        /// Konvertiert eine zweidimensionale Zeichenketten-Matrix (<see cref="string[,]"/>) 
        /// in eine <see cref="DataView"/>, sodass sie z. B. in einem <see cref="System.Windows.Controls.DataGrid"/> angezeigt werden kann.
        /// </summary>
        /// <param name="matrix">
        /// Zweidimensionale Matrix mit Textwerten. 
        /// Wenn <c>null</c> oder leer, wird eine leere <see cref="DataView"/> zurückgegeben.
        /// </param>
        /// <param name="tableName">
        /// Name der zu erstellenden <see cref="DataTable"/> (z. B. zur Identifikation im Debug oder Binding).
        /// </param>
        /// <returns>
        /// Eine <see cref="DataView"/> mit derselben Struktur und denselben Werten wie die Eingabematrix.
        /// </returns>
        public static DataView ToDataView(string[,]? matrix, string tableName)
        {
            var dt = new DataTable(tableName);

            if (matrix is null || matrix.GetLength(0) == 0 || matrix.GetLength(1) == 0)
                return dt.DefaultView;

            int rows = matrix.GetLength(0);
            int cols = matrix.GetLength(1);

            for (int c = 0; c < cols; c++)
                dt.Columns.Add($"Col{c}", typeof(string));

            for (int r = 0; r < rows; r++)
            {
                var row = dt.NewRow();
                for (int c = 0; c < cols; c++)
                    row[c] = matrix[r, c] ?? string.Empty;
                dt.Rows.Add(row);
            }

            return dt.DefaultView;
        }

        /// <summary>
        /// Konvertiert eine zweidimensionale Ganzzahl-Matrix (<see cref="int[,]"/>) 
        /// in eine <see cref="DataView"/> zur tabellarischen Anzeige oder Weiterverarbeitung.
        /// </summary>
        /// <param name="matrix">
        /// Zweidimensionale Matrix mit Ganzzahlen. 
        /// Wenn <c>null</c> oder leer, wird eine leere <see cref="DataView"/> zurückgegeben.
        /// </param>
        /// <param name="tableName">
        /// Name der zu erstellenden <see cref="DataTable"/> (z. B. zur Identifikation im Debug oder Binding).
        /// </param>
        /// <returns>
        /// Eine <see cref="DataView"/> mit derselben Struktur und denselben Werten wie die Eingabematrix.
        /// </returns>
        public static DataView ToDataView(int[,]? matrix, string tableName)
        {
            var dt = new DataTable(tableName);

            if (matrix is null || matrix.GetLength(0) == 0 || matrix.GetLength(1) == 0)
                return dt.DefaultView;

            int rows = matrix.GetLength(0);
            int cols = matrix.GetLength(1);

            for (int c = 0; c < cols; c++)
                dt.Columns.Add($"Col{c}", typeof(int));

            for (int r = 0; r < rows; r++)
            {
                var row = dt.NewRow();
                for (int c = 0; c < cols; c++)
                    row[c] = matrix[r, c];
                dt.Rows.Add(row);
            }

            return dt.DefaultView;
        }
    }
}
