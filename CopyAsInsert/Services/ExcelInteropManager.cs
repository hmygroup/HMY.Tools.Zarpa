using System.Runtime.InteropServices;
using System.Reflection;
using CopyAsInsert.Models;
// Note: keep logic so the query is executed by Excel via QueryTables

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
                Logger.LogError("Excel.Application ProgID not found");
                throw new Exception("Excel no está instalado o no se puede acceder via COM");
            }
            Logger.LogDebug($"Resolved Excel.Application to type '{excelType.FullName}'");

            // Crear instancia de Excel COM object
            excelApp = Activator.CreateInstance(excelType);
            if (excelApp == null)
            {
                Logger.LogError("Activator.CreateInstance returned null for Excel.Application");
                throw new Exception("No se pudo crear instancia de Excel");
            }
            Logger.LogInfo("Excel COM instance created");
            try
            {
                var ver = excelApp.GetType().InvokeMember("Version", BindingFlags.GetProperty, null, excelApp, null);
                Logger.LogDebug($"Excel version: {ver}");
            }
            catch (Exception ex)
            {
                Logger.LogDebug($"Could not read Excel version: {ex.Message}");
            }

            // Set Visible = true using reflection
            excelApp.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, excelApp, new object[] { true });
            Logger.LogDebug("Set Excel.Visible = true");
            
            // Crear workbook nuevo: excelApp.Workbooks.Add()
            object? workbooks = excelApp.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, excelApp, null);
            workbook = workbooks?.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, workbooks, null);
            if (workbook == null)
            {
                Logger.LogError("Failed to create new workbook via Excel.Workbooks.Add");
                throw new Exception("No se pudo crear workbook");
            }
            Logger.LogDebug("Workbook created");

            // Get Sheets[1] (first worksheet)
            object? sheets = workbook.GetType().InvokeMember("Sheets", BindingFlags.GetProperty, null, workbook, null);
            object? worksheet = sheets?.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, sheets, new object[] { 1 });
            if (worksheet == null)
            {
                throw new Exception("No se pudo obtener worksheet");
            }

            // Get Range["A1"]
            object? destination = worksheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, worksheet,
                new object[] { "A1" });
            Logger.LogDebug("Obtained destination Range A1");

            // Try to use Excel QueryTables with common providers. If all providers fail, throw a descriptive error
            // so that Excel remains responsible for executing the query (do not run the SQL locally).
            object? queryTables = null;
            try
            {
                queryTables = worksheet.GetType().InvokeMember("QueryTables", BindingFlags.GetProperty, null, worksheet, null);
                Logger.LogDebug("QueryTables collection obtained from worksheet");
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"QueryTables not available: {ex.Message}");
                Logger.LogDebug(ex.ToString());
                queryTables = null;
            }

            object? qt = null;
            Exception? lastProviderException = null;

            if (queryTables != null)
            {
                string[] providers = new[] { "MSOLEDBSQL", "SQLNCLI11", "SQLOLEDB" };
                Logger.LogDebug($"Attempting QueryTable.Add using providers: {string.Join(',', providers)}");
                foreach (var prov in providers)
                {
                    try
                    {
                        string connectionString = $"OLEDB;Provider={prov};Server={server};Database={database};Integrated Security=SSPI;Persist Security Info=False;";
                        Logger.LogDebug($"Trying provider {prov} with connection string: {connectionString}");
                        try
                        {
                            qt = queryTables.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, queryTables,
                                new object[] { connectionString, destination, query });
                        }
                        catch (TargetInvocationException tie)
                        {
                            lastProviderException = tie.InnerException ?? tie;
                            Logger.LogWarning($"Provider {prov} Add() failed: {lastProviderException.Message}");
                            Logger.LogDebug(lastProviderException.ToString());
                            qt = null;
                        }

                        if (qt != null)
                        {
                            Logger.LogInfo($"QueryTable created using provider {prov}");
                            // Configure and refresh
                            qt.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, qt, new object[] { $"Query_{DateTime.Now:yyyyMMdd_HHmmss}" });
                            qt.GetType().InvokeMember("FieldNames", BindingFlags.SetProperty, null, qt, new object[] { true });
                            qt.GetType().InvokeMember("RowNumbers", BindingFlags.SetProperty, null, qt, new object[] { false });
                            qt.GetType().InvokeMember("FillAdjacentFormulas", BindingFlags.SetProperty, null, qt, new object[] { false });
                            qt.GetType().InvokeMember("PreserveFormatting", BindingFlags.SetProperty, null, qt, new object[] { true });
                            qt.GetType().InvokeMember("RefreshStyle", BindingFlags.SetProperty, null, qt, new object[] { 0 });
                            qt.GetType().InvokeMember("SavePassword", BindingFlags.SetProperty, null, qt, new object[] { false });
                            qt.GetType().InvokeMember("SaveData", BindingFlags.SetProperty, null, qt, new object[] { true });
                            qt.GetType().InvokeMember("AdjustColumnWidth", BindingFlags.SetProperty, null, qt, new object[] { true });
                            try
                            {
                                qt.GetType().InvokeMember("Refresh", BindingFlags.InvokeMethod, null, qt, new object[] { false });
                                Logger.LogDebug("QueryTable.Refresh() invoked");
                            }
                            catch (TargetInvocationException tie)
                            {
                                var ie = tie.InnerException ?? tie;
                                Logger.LogWarning($"Refresh failed for provider {prov}: {ie.Message}");
                                Logger.LogDebug(ie.ToString());
                            }
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        lastProviderException = ex;
                        Logger.LogWarning($"Provider {prov} test failed: {ex.Message}");
                        Logger.LogDebug(ex.ToString());
                        qt = null;
                    }
                }
            }

            if (qt == null)
            {
                Logger.LogError("All tested providers failed to create a QueryTable.");
                if (lastProviderException != null)
                {
                    Logger.LogError(lastProviderException.ToString());
                }
                throw new Exception("QueryTable.Add failed for all tested providers. See application log for details.");
            }
            
            int rowCount = 0;
            try
            {
                if (qt != null)
                {
                    object? resultRange = qt.GetType().InvokeMember("ResultRange", BindingFlags.GetProperty, null, qt, null);
                    if (resultRange != null)
                    {
                        object? rows = resultRange.GetType().InvokeMember("Rows", BindingFlags.GetProperty, null, resultRange, null);
                        if (rows != null)
                        {
                            object? count = rows.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, rows, null);
                            if (count is int intCount && intCount > 1)
                                rowCount = intCount - 1;
                            Logger.LogDebug($"ResultRange rows detected: {rowCount}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"Error reading ResultRange: {ex.Message}");
                Logger.LogDebug(ex.ToString());
            }
            
            // Traer Excel al frente
            try
            {
                object? hwnd = excelApp.GetType().InvokeMember("Hwnd", BindingFlags.GetProperty, null, excelApp, null);
                Logger.LogDebug($"Excel Hwnd property: {hwnd}");
                if (hwnd is int hWndInt)
                {
                    SetForegroundWindow((IntPtr)hWndInt);
                    Logger.LogDebug("Set Excel window to foreground");
                }
            }
            catch (Exception ex)
            {
                Logger.LogDebug($"Could not bring Excel to front: {ex.Message}");
            }
            
            // Get workbook full name
            string outputPath = "";
            try
            {
                object? fullName = workbook.GetType().InvokeMember("FullName", BindingFlags.GetProperty, null, workbook, null);
                outputPath = fullName?.ToString() ?? "";
                Logger.LogDebug($"Workbook FullName: {outputPath}");
            }
            catch (Exception ex)
            {
                Logger.LogDebug($"Could not read workbook FullName: {ex.Message}");
            }
            
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
            Logger.LogDebug(ex.ToString());

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
            catch (Exception closeEx)
            {
                Logger.LogDebug($"Error closing Excel after failure: {closeEx.Message}");
            }

            return new ImportResult
            {
                Success = false,
                RowCount = 0,
                ErrorMessage = $"Error al inyectar query: {ex.Message}",
                ErrorStackTrace = ex.ToString(),
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
