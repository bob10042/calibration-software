using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace CalTestGUI.Models
{
    /// <summary>
    /// Provides helper methods for safely working with Excel worksheets and cell values
    /// </summary>
    public static class ExcelHelpers
    {
        /// <summary>
        /// Safely retrieves a cell value as a double from an Excel worksheet
        /// </summary>
        /// <param name="worksheet">The Excel worksheet</param>
        /// <param name="row">The row number (1-based)</param>
        /// <param name="column">The column number (1-based)</param>
        /// <param name="defaultValue">The default value to return if conversion fails</param>
        /// <returns>The cell value as a double, or the default value</returns>
        public static double SafeGetCellValueAsDouble(
            Excel.Worksheet worksheet, 
            int row, 
            int column, 
            double defaultValue = 0)
        {
            try
            {
                // Retrieve the cell value
                object cellValue = (worksheet.Cells[row, column] as Excel.Range)?.Value;
                
                // Check for null or DBNull
                if (cellValue == null || cellValue == DBNull.Value)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"Warning: Null cell value at [{row},{column}]. Using default.");
                    return defaultValue;
                }

                // If it's already a double, return it
                if (cellValue is double doubleValue)
                {
                    return doubleValue;
                }

                // Try parsing as a string
                if (double.TryParse(cellValue.ToString(), out double parsedValue))
                {
                    return parsedValue;
                }

                // Log warning and return default
                System.Diagnostics.Debug.WriteLine(
                    $"Warning: Could not convert cell value at [{row},{column}] to double. Value: {cellValue}");
                return defaultValue;
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                System.Diagnostics.Debug.WriteLine(
                    $"Error converting cell value at [{row},{column}]: {ex.Message}");
                return defaultValue;
            }
        }

        /// <summary>
        /// Safely retrieves a cell value as a string from an Excel worksheet
        /// </summary>
        /// <param name="worksheet">The Excel worksheet</param>
        /// <param name="row">The row number (1-based)</param>
        /// <param name="column">The column number (1-based)</param>
        /// <param name="defaultValue">The default value to return if conversion fails</param>
        /// <returns>The cell value as a string, or the default value</returns>
        public static string SafeGetCellValueAsString(
            Excel.Worksheet worksheet, 
            int row, 
            int column, 
            string defaultValue = "")
        {
            try
            {
                // Retrieve the cell value
                object cellValue = (worksheet.Cells[row, column] as Excel.Range)?.Value;
                
                // Check for null or DBNull
                if (cellValue == null || cellValue == DBNull.Value)
                {
                    return defaultValue;
                }

                // Convert to string
                return cellValue.ToString().Trim();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"Error converting cell value at [{row},{column}]: {ex.Message}");
                return defaultValue;
            }
        }
    }
}
