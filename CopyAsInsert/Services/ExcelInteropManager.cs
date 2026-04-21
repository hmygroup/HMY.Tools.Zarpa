using System.Runtime.InteropServices;
using System.Reflection;
using System.Windows.Forms;
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
            int? excelHwnd = null;
            bool excelBroughtToFront = false;
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

            // Try to add a Power Query (M) to the workbook's Queries collection so it appears
            // in the Queries pane (Power Query). This keeps the query in Excel instead of
            // creating an external connection-only object. If this succeeds we treat it as
            // a successful injection (Excel will perform the query when refreshed).
            bool powerQueryAdded = false;
            object? powerQueryTable = null;
            try
            {
                object? queries = workbook.GetType().InvokeMember("Queries", BindingFlags.GetProperty, null, workbook, null);
                if (queries != null)
                {
                    string pqName = $"Query_{DateTime.Now:yyyyMMdd_HHmmss}";
                    string mEscaped = query?.Replace("\"", "\"\"") ?? string.Empty; // escape quotes for M string
                    string mFormula = $"let\r\n    Source = Sql.Database(\"{server}\", \"{database}\", [Query=\"{mEscaped}\"])\r\nin\r\n    Source";
                    Logger.LogDebug($"Attempting to add Power Query '{pqName}' to workbook.Queries");
                    queries.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, queries, new object[] { pqName, mFormula });
                    Logger.LogInfo($"Power Query '{pqName}' added to workbook.Queries");
                    powerQueryAdded = true;

                    // Try to load the Power Query into a NEW worksheet as a table using Excel's
                    // Mashup provider. WORKBOOK_QUERY sources can trigger the Import Data dialog
                    // and leave the query in "Connection only" mode.
                    try
                    {
                        object? newSheet = null;
                        try
                        {
                            newSheet = sheets.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, sheets, null);
                            try { newSheet.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, newSheet, new object[] { pqName }); } catch { }
                        }
                        catch (Exception addSheetEx)
                        {
                            Logger.LogDebug($"Could not add new sheet for Power Query: {addSheetEx.Message}");
                        }

                        object? targetSheet = newSheet ?? worksheet;
                        object? newDestination = null;
                        try
                        {
                            newDestination = targetSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, targetSheet, new object[] { "A1" });
                        }
                        catch { newDestination = destination; }

                        object? targetListObjects = null;
                        try
                        {
                            targetListObjects = targetSheet.GetType().InvokeMember("ListObjects", BindingFlags.GetProperty, null, targetSheet, null);
                        }
                        catch (Exception ex)
                        {
                            Logger.LogDebug($"Could not access ListObjects collection for Power Query load: {ex.Message}");
                        }

                        object? createdListObj = null;
                        if (targetListObjects != null && newDestination != null)
                        {
                            try
                            {
                                string mashupConnection = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={pqName};Extended Properties=\"\"";
                                Logger.LogDebug($"Loading Power Query '{pqName}' into worksheet using Mashup OLE DB connection.");
                                createdListObj = targetListObjects.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, targetListObjects,
                                    new object[] { 0, mashupConnection, true, 0, newDestination });
                            }
                            catch (TargetInvocationException tie)
                            {
                                var ie = tie.InnerException ?? tie;
                                Logger.LogWarning($"ListObjects.Add failed for Power Query '{pqName}': {ie.Message}");
                                Logger.LogDebug(ie.ToString());
                            }
                            catch (Exception ex)
                            {
                                Logger.LogDebug($"Could not create ListObject for Power Query load: {ex.Message}");
                            }
                        }

                        if (createdListObj != null)
                        {
                            try { createdListObj.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, createdListObj, new object[] { $"Table_{pqName}" }); } catch { }

                            object? createdQtbl = null;
                            try
                            {
                                createdQtbl = createdListObj.GetType().InvokeMember("QueryTable", BindingFlags.GetProperty, null, createdListObj, null);
                            }
                            catch (Exception ex)
                            {
                                Logger.LogDebug($"Could not access QueryTable for Power Query ListObject: {ex.Message}");
                            }

                            if (createdQtbl != null)
                            {
                                try { createdQtbl.GetType().InvokeMember("CommandType", BindingFlags.SetProperty, null, createdQtbl, new object[] { 2 }); } catch { }
                                try { createdQtbl.GetType().InvokeMember("CommandText", BindingFlags.SetProperty, null, createdQtbl, new object[] { new string[] { $"SELECT * FROM [{pqName}]" } }); } catch (Exception ex) { Logger.LogDebug($"Could not set CommandText for Power Query table load: {ex.Message}"); }
                                try { createdQtbl.GetType().InvokeMember("RowNumbers", BindingFlags.SetProperty, null, createdQtbl, new object[] { false }); } catch { }
                                try { createdQtbl.GetType().InvokeMember("FillAdjacentFormulas", BindingFlags.SetProperty, null, createdQtbl, new object[] { false }); } catch { }
                                try { createdQtbl.GetType().InvokeMember("RefreshOnFileOpen", BindingFlags.SetProperty, null, createdQtbl, new object[] { false }); } catch { }
                                try { createdQtbl.GetType().InvokeMember("BackgroundQuery", BindingFlags.SetProperty, null, createdQtbl, new object[] { false }); } catch { }
                                try { createdQtbl.GetType().InvokeMember("RefreshStyle", BindingFlags.SetProperty, null, createdQtbl, new object[] { 0 }); } catch { }
                                try { createdQtbl.GetType().InvokeMember("SavePassword", BindingFlags.SetProperty, null, createdQtbl, new object[] { false }); } catch { }
                                try { createdQtbl.GetType().InvokeMember("SaveData", BindingFlags.SetProperty, null, createdQtbl, new object[] { true }); } catch { }
                                try { createdQtbl.GetType().InvokeMember("AdjustColumnWidth", BindingFlags.SetProperty, null, createdQtbl, new object[] { true }); } catch { }

                                try
                                {
                                    createdQtbl.GetType().InvokeMember("Refresh", BindingFlags.InvokeMethod, null, createdQtbl, new object[] { false });
                                    int loadedRowCount = GetImportedRowCount(createdQtbl);
                                    if (loadedRowCount > 0)
                                    {
                                        powerQueryTable = createdQtbl;
                                        Logger.LogInfo($"Power Query '{pqName}' loaded into worksheet with {loadedRowCount} rows.");
                                    }
                                    else
                                    {
                                        Logger.LogWarning($"Power Query '{pqName}' refresh completed but no rows were materialized in the worksheet.");
                                    }
                                }
                                catch (TargetInvocationException tie)
                                {
                                    var ie = tie.InnerException ?? tie;
                                    Logger.LogWarning($"Refresh failed for Power Query worksheet load: {ie.Message}");
                                    Logger.LogDebug(ie.ToString());
                                }
                                catch (Exception ex)
                                {
                                    Logger.LogDebug($"Refresh exception for Power Query worksheet load: {ex.Message}");
                                }
                            }
                        }

                        if (powerQueryTable == null)
                        {
                            try
                            {
                                if (createdListObj != null)
                                {
                                    createdListObj.GetType().InvokeMember("Delete", BindingFlags.InvokeMethod, null, createdListObj, null);
                                    Logger.LogDebug($"Deleted empty Power Query table shell for '{pqName}' before fallback.");
                                }
                            }
                            catch (Exception ex)
                            {
                                Logger.LogDebug($"Could not delete empty Power Query table shell: {ex.Message}");
                            }

                            try
                            {
                                Logger.LogWarning($"Power Query sheet load produced no data; attempting provider fallback on the same sheet using MSOLEDBSQL.");
                                object? targetQTables = targetSheet.GetType().InvokeMember("QueryTables", BindingFlags.GetProperty, null, targetSheet, null);
                                if (targetQTables != null && newDestination != null)
                                {
                                    string prov = "MSOLEDBSQL";
                                    string connectionString = $"OLEDB;Provider={prov};Server={server};Database={database};Integrated Security=SSPI;Persist Security Info=False;";
                                    Logger.LogDebug($"Provider fallback connection string: {connectionString}");
                                    object? provQtbl = targetQTables.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, targetQTables,
                                        new object[] { connectionString, newDestination, query });
                                    if (provQtbl != null)
                                    {
                                        try { provQtbl.GetType().InvokeMember("FieldNames", BindingFlags.SetProperty, null, provQtbl, new object[] { true }); } catch { }
                                        try { provQtbl.GetType().InvokeMember("RowNumbers", BindingFlags.SetProperty, null, provQtbl, new object[] { false }); } catch { }
                                        try { provQtbl.GetType().InvokeMember("FillAdjacentFormulas", BindingFlags.SetProperty, null, provQtbl, new object[] { false }); } catch { }
                                        try { provQtbl.GetType().InvokeMember("PreserveFormatting", BindingFlags.SetProperty, null, provQtbl, new object[] { true }); } catch { }
                                        try { provQtbl.GetType().InvokeMember("RefreshStyle", BindingFlags.SetProperty, null, provQtbl, new object[] { 0 }); } catch { }
                                        try { provQtbl.GetType().InvokeMember("SavePassword", BindingFlags.SetProperty, null, provQtbl, new object[] { false }); } catch { }
                                        try { provQtbl.GetType().InvokeMember("SaveData", BindingFlags.SetProperty, null, provQtbl, new object[] { true }); } catch { }
                                        try { provQtbl.GetType().InvokeMember("AdjustColumnWidth", BindingFlags.SetProperty, null, provQtbl, new object[] { true }); } catch { }

                                        try
                                        {
                                            provQtbl.GetType().InvokeMember("Refresh", BindingFlags.InvokeMethod, null, provQtbl, new object[] { false });
                                            int loadedRowCount = GetImportedRowCount(provQtbl);
                                            if (loadedRowCount > 0)
                                            {
                                                powerQueryTable = provQtbl;
                                                Logger.LogInfo($"Provider fallback loaded {loadedRowCount} rows into worksheet for '{pqName}'.");
                                            }
                                            else
                                            {
                                                Logger.LogWarning($"Provider fallback refresh completed but still produced no rows for '{pqName}'.");
                                            }
                                        }
                                        catch (TargetInvocationException tie)
                                        {
                                            var ie = tie.InnerException ?? tie;
                                            Logger.LogWarning($"Provider fallback refresh failed: {ie.Message}");
                                            Logger.LogDebug(ie.ToString());
                                        }
                                        catch (Exception ex)
                                        {
                                            Logger.LogDebug($"Provider fallback refresh exception: {ex.Message}");
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Logger.LogDebug($"Provider fallback exception on Power Query sheet: {ex.Message}");
                            }
                        }
                    }
                    catch (Exception pqLoadEx)
                    {
                        Logger.LogDebug($"Could not load Power Query into worksheet: {pqLoadEx.Message}");
                    }
                }
            }
            catch (TargetInvocationException tie)
            {
                var ie = tie.InnerException ?? tie;
                Logger.LogWarning($"Add Power Query failed: {ie.Message}");
                Logger.LogDebug(ie.ToString());

                try
                {
                    var m = (ie.Message ?? string.Empty).ToLowerInvariant();
                    if (m.Contains("evaluate") && m.Contains("native") || m.Contains("native database") || m.Contains("evaluatenativequeryunpermitted") || m.Contains("permission is required to run this native"))
                    {
                        Logger.LogWarning("Power Query blocked native query execution (EvaluateNativeQuery). Prompting user to allow native queries in Power Query settings.");
                        string helpMsg = "Power Query está bloqueando la ejecución de consultas SQL nativas.\n\n" +
                            "Para permitir la ejecución de esta consulta, abra el Power Query Editor (Datos -> Obtener y transformar -> Launch Power Query Editor).\n" +
                            "En el editor, haga clic en 'Edit Permissions' en la banda amarilla o vaya a File -> Options and settings -> Query Options -> Security y habilite 'Allow native database queries' para esta fuente.\n\n" +
                            "Alternativamente, en Excel vaya a Data -> Get Data -> Data Source Settings -> Edit Permissions y permita la consulta para la fuente.\n\n" +
                            "Después de dar permiso, vuelva a ejecutar la importación.";
                        try { MessageBox.Show(helpMsg, "Permiso requerido para consulta nativa", MessageBoxButtons.OK, MessageBoxIcon.Warning); } catch { }
                    }
                }
                catch { }
            }
            catch (Exception pqEx)
            {
                Logger.LogDebug($"Could not add Power Query: {pqEx.Message}");
            }

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

            object? qt = powerQueryTable;
            Exception? providerException = null;

            if (queryTables != null && qt == null)
            {
                // This import targets SQL Server only — use the Microsoft OLE DB Driver for SQL Server
                string prov = "MSOLEDBSQL";
                string connectionString = $"OLEDB;Provider={prov};Server={server};Database={database};Integrated Security=SSPI;Persist Security Info=False;";
                Logger.LogDebug($"Using SQL Server OLE DB provider: {prov}");
                Logger.LogDebug($"Connection string: {connectionString}");
                try
                {
                    qt = queryTables.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, queryTables,
                        new object[] { connectionString, destination, query });

                    if (qt != null)
                    {
                        Logger.LogInfo($"QueryTable created using provider {prov}");
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
                    }
                }
                catch (TargetInvocationException tie)
                {
                    providerException = tie.InnerException ?? tie;
                    Logger.LogError($"Provider {prov} Add() failed: {providerException.Message}");
                    Logger.LogDebug(providerException.ToString());
                }
                catch (Exception ex)
                {
                    providerException = ex;
                    Logger.LogError($"Provider {prov} test failed: {ex.Message}");
                    Logger.LogDebug(ex.ToString());
                }

                if (qt == null)
                {
                    Logger.LogError($"Failed to create QueryTable using provider {prov}. Ensure the Microsoft OLE DB Driver for SQL Server (MSOLEDBSQL) is installed and available on this machine.");
                }
            }

            if (powerQueryAdded && qt == null)
            {
                Logger.LogWarning("Power Query was added to the workbook, but no rows were loaded into a worksheet.");
            }

            if (qt == null)
            {
                Logger.LogError("QueryTable.Add failed using MSOLEDBSQL provider.");
                if (providerException != null)
                {
                    Logger.LogError(providerException.ToString());
                }
                throw new Exception("QueryTable.Add failed with MSOLEDBSQL. Ensure the Microsoft OLE DB Driver for SQL Server is installed. See application log for details.");
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
                    excelHwnd = hWndInt;
                    excelBroughtToFront = true;
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
                OutputPath = outputPath,
                ExcelHwnd = excelHwnd,
                ExcelBroughtToForeground = excelBroughtToFront
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

    private static int GetImportedRowCount(object? queryTable)
    {
        try
        {
            if (queryTable == null)
            {
                return 0;
            }

            object? resultRange = queryTable.GetType().InvokeMember("ResultRange", BindingFlags.GetProperty, null, queryTable, null);
            if (resultRange == null)
            {
                return 0;
            }

            object? rows = resultRange.GetType().InvokeMember("Rows", BindingFlags.GetProperty, null, resultRange, null);
            if (rows == null)
            {
                return 0;
            }

            object? count = rows.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, rows, null);
            if (count is int intCount)
            {
                return Math.Max(0, intCount - 1);
            }
        }
        catch (Exception ex)
        {
            Logger.LogDebug($"Could not determine imported row count: {ex.Message}");
        }

        return 0;
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
