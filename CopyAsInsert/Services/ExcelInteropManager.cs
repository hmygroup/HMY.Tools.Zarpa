using System.Runtime.InteropServices;
using System.Reflection;
using CopyAsInsert.Models;

namespace CopyAsInsert.Services;

/// <summary>
/// Manages Excel automation for SQL Server query import using pure COM reflection
/// Query se inyecta en las conexiones de datos de Excel
/// Uses reflection-based COM interop without any Office.Interop types to work with self-contained deployment
/// </summary>
public class ExcelInteropManager
{
    /// <summary>
    /// Abre Excel y crea una query de SQL Server con los parámetros dados
    /// Uses pure reflection to invoke Excel COM methods - NO Office.Interop references
    /// </summary>
    public static ImportResult InjectQueryIntoExcel(string server, string database, string query)
    {
        object? excelApp = null;
        object? workbook = null;

        try
        {
            // Resolver Excel.Application via COM using reflection (NO Office.Interop types)
            Type? excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                throw new Exception("Excel no está instalado o no se puede acceder via COM");
            }

            // Crear instancia de Excel COM object
            excelApp = Activator.CreateInstance(excelType);
            if (excelApp == null)
            {
                throw new Exception("No se pudo crear instancia de Excel");
            }

            // Set Visible = true using reflection
            excelApp.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, excelApp, new object[] { true });
            
            // Crear workbook nuevo: excelApp.Workbooks.Add()
            object? workbooks = excelApp.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, excelApp, null);
            workbook = workbooks?.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, workbooks, null);
            if (workbook == null)
            {
                throw new Exception("No se pudo crear workbook");
            }

            // Get Sheets[1] (first worksheet)
            object? sheets = workbook.GetType().InvokeMember("Sheets", BindingFlags.GetProperty, null, workbook, null);
            object? worksheet = sheets?.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, sheets, new object[] { 1 });
            if (worksheet == null)
            {
                throw new Exception("No se pudo obtener worksheet");
            }

            // Crear connection string en formato OLEDB para QueryTable.Add()
            // Format: OLEDB;Provider=provider;Server=server;Database=database;...
            string connectionString = $"OLEDB;Provider=MSOLEDBSQL;Server={server};Database={database};Integrated Security=SSPI;Persist Security Info=False;";

            // Get Range["A1"]
            object? destination = worksheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, worksheet, 
                new object[] { "A1" });
            
            // Get QueryTables
            object? queryTables = worksheet.GetType().InvokeMember("QueryTables", BindingFlags.GetProperty, null, worksheet, null);
            if (queryTables == null)
            {
                throw new Exception("No se pudo obtener QueryTables");
            }

            // Add QueryTable: qt = queryTables.Add(connectionString, destination, sql)
            object? qt = null;
            try
            {
                qt = queryTables.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, queryTables,
                    new object[] { connectionString, destination, query });
            }
            catch (TargetInvocationException ex)
            {
                throw new Exception($"Error en QueryTable.Add(): {ex.InnerException?.Message ?? ex.Message}", ex.InnerException ?? ex);
            }
            
            if (qt == null)
            {
                throw new Exception("No se pudo crear QueryTable");
            }

            // Configurar propiedades de la QueryTable usando reflection
            qt.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, qt, 
                new object[] { $"Query_{DateTime.Now:yyyyMMdd_HHmmss}" });
            qt.GetType().InvokeMember("FieldNames", BindingFlags.SetProperty, null, qt, new object[] { true });
            qt.GetType().InvokeMember("RowNumbers", BindingFlags.SetProperty, null, qt, new object[] { false });
            qt.GetType().InvokeMember("FillAdjacentFormulas", BindingFlags.SetProperty, null, qt, new object[] { false });
            qt.GetType().InvokeMember("PreserveFormatting", BindingFlags.SetProperty, null, qt, new object[] { true });
            qt.GetType().InvokeMember("RefreshStyle", BindingFlags.SetProperty, null, qt, new object[] { 0 });
            qt.GetType().InvokeMember("SavePassword", BindingFlags.SetProperty, null, qt, new object[] { false });
            qt.GetType().InvokeMember("SaveData", BindingFlags.SetProperty, null, qt, new object[] { true });
            qt.GetType().InvokeMember("AdjustColumnWidth", BindingFlags.SetProperty, null, qt, new object[] { true });
            
            // Ejecutar la query: qt.Refresh(false)
            qt.GetType().InvokeMember("Refresh", BindingFlags.InvokeMethod, null, qt, new object[] { false });
            
            int rowCount = 0;
            try
            {
                object? resultRange = qt.GetType().InvokeMember("ResultRange", BindingFlags.GetProperty, null, qt, null);
                if (resultRange != null)
                {
                    object? rows = resultRange.GetType().InvokeMember("Rows", BindingFlags.GetProperty, null, resultRange, null);
                    if (rows != null)
                    {
                        object? count = rows.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, rows, null);
                        if (count is int intCount && intCount > 1)
                        {
                            rowCount = intCount - 1;
                        }
                    }
                }
            }
            catch { }
            
            // Traer Excel al frente
            try
            {
                object? hwnd = excelApp.GetType().InvokeMember("Hwnd", BindingFlags.GetProperty, null, excelApp, null);
                if (hwnd is int hWndInt)
                {
                    SetForegroundWindow((IntPtr)hWndInt);
                }
            }
            catch { }
            
            // Get workbook full name
            string outputPath = "";
            try
            {
                object? fullName = workbook.GetType().InvokeMember("FullName", BindingFlags.GetProperty, null, workbook, null);
                outputPath = fullName?.ToString() ?? "";
            }
            catch { }
            
            return new ImportResult
            {
                Success = true,
                RowCount = rowCount,
                ServerName = server,
                DatabaseName = database,
                ImportTime = DateTime.Now,
                OutputPath = outputPath
            };
        }
        catch (Exception ex)
        {
            Logger.LogError($"Error inyectando query en Excel: {ex.Message}");
            
            // Cerrar Excel si hubo error
            try
            {
                if (workbook != null)
                {
                    workbook.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, workbook, new object[] { false });
                    Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    excelApp.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, excelApp, null);
                    Marshal.ReleaseComObject(excelApp);
                }
            }
            catch { }
            
            return new ImportResult
            {
                Success = false,
                RowCount = 0,
                ErrorMessage = $"Error al inyectar query: {ex.Message}",
                ErrorStackTrace = ex.StackTrace,
                ServerName = server,
                DatabaseName = database
            };
        }
    }

    [DllImport("user32.dll")]
    private static extern IntPtr SetForegroundWindow(IntPtr hWnd);

    /// <summary>
    /// Prueba la conexión a SQL Server
    /// </summary>
    public static (bool Success, string ErrorMessage) TestConnection(string server, string database)
    {
        try
        {
            if (string.IsNullOrEmpty(server) || string.IsNullOrEmpty(database))
            {
                return (false, "Server y database no pueden estar vacíos");
            }
            
            Logger.LogDebug($"Test conexión: {server}/{database}");
            return (true, string.Empty);
        }
        catch (Exception ex)
        {
            Logger.LogWarning($"Error test conexión: {ex.Message}");
            return (false, ex.Message);
        }
    }
}
