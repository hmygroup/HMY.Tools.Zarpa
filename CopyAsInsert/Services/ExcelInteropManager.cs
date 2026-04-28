using System.Runtime.InteropServices;
using System.Reflection;
using System.Windows.Forms;
using System.Collections.Generic;
using CopyAsInsert.Models;
// Note: keep logic so the query is executed by Excel via Power Query

namespace CopyAsInsert.Services;

/// <summary>
/// Manages Excel automation for SQL Server query import using pure COM reflection
/// Query se inyecta en las conexiones de datos de Excel
/// Uses reflection-based COM interop without any Office.Interop types to work with self-contained deployment
/// </summary>
public class ExcelInteropManager
{
    public sealed class OpenWorkbookInfo
    {
        public string Name { get; init; } = string.Empty;
        public string FullName { get; init; } = string.Empty;
        public List<string> WorksheetNames { get; init; } = new();
        public string WorkbookKey => string.IsNullOrWhiteSpace(FullName) ? Name : FullName;
        public string DisplayName => string.IsNullOrWhiteSpace(FullName) ? Name : $"{Name} ({FullName})";

        public override string ToString() => DisplayName;
    }

    public sealed class ImportTargetOptions
    {
        public bool UseOpenWorkbook { get; set; }
        public string? WorkbookKey { get; set; }
        public string? WorkbookName { get; set; }
        public string? WorksheetName { get; set; }
        public bool CreateNewWorksheet { get; set; }
    }

    public static List<OpenWorkbookInfo> GetOpenWorkbooks()
    {
        object? excelApp = null;
        object? workbooks = null;

        try
        {
            excelApp = TryGetRunningExcelApplication();
            if (excelApp == null)
            {
                Logger.LogDebug("No running Excel instance found while listing open workbooks.");
                return new List<OpenWorkbookInfo>();
            }

            workbooks = excelApp.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, excelApp, null);
            if (workbooks == null)
            {
                return new List<OpenWorkbookInfo>();
            }

            int workbookCount = GetComCount(workbooks);
            List<OpenWorkbookInfo> openWorkbooks = new();

            for (int index = 1; index <= workbookCount; index++)
            {
                object? workbook = null;
                try
                {
                    workbook = workbooks.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, workbooks, new object[] { index });
                    if (workbook == null)
                    {
                        continue;
                    }

                    string workbookName = GetComStringProperty(workbook, "Name") ?? $"Workbook{index}";
                    string workbookFullName = GetComStringProperty(workbook, "FullName") ?? workbookName;
                    List<string> worksheetNames = GetWorksheetNames(workbook);

                    openWorkbooks.Add(new OpenWorkbookInfo
                    {
                        Name = workbookName,
                        FullName = workbookFullName,
                        WorksheetNames = worksheetNames
                    });
                }
                catch (Exception ex)
                {
                    Logger.LogDebug($"Could not inspect open workbook at index {index}: {ex.Message}");
                }
                finally
                {
                    SafeReleaseComObject(workbook);
                }
            }

            return openWorkbooks;
        }
        catch (Exception ex)
        {
            Logger.LogWarning($"Could not list open Excel workbooks: {ex.Message}");
            return new List<OpenWorkbookInfo>();
        }
        finally
        {
            SafeReleaseComObject(workbooks);
            SafeReleaseComObject(excelApp);
        }
    }

    /// <summary>
    /// Abre Excel y crea una query de SQL Server con los parámetros dados
    /// Uses pure reflection to invoke Excel COM methods - NO Office.Interop references
    /// </summary>
    public static ImportResult InjectQueryIntoExcel(string server, string database, string query, ImportTargetOptions? targetOptions = null)
    {
        object? excelApp = null;
        object? workbook = null;
        object? sharedWorksheet = null;
        bool attachedToRunningExcel = false;
        bool createdWorkbook = false;
        bool useOpenWorkbook = targetOptions?.UseOpenWorkbook == true;

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

            if (useOpenWorkbook)
            {
                excelApp = TryGetRunningExcelApplication();
                attachedToRunningExcel = excelApp != null;
                if (excelApp == null)
                {
                    throw new Exception("No hay ninguna instancia de Excel abierta. Desmarca la opción para usar el flujo normal o abre un libro antes de importar.");
                }

                Logger.LogInfo("Connected to running Excel instance");
            }
            else
            {
                // Crear instancia de Excel COM object
                excelApp = Activator.CreateInstance(excelType);
            }

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
            
            object? workbooks = excelApp.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, excelApp, null);
            if (useOpenWorkbook)
            {
                workbook = FindWorkbook(workbooks, targetOptions);
                if (workbook == null)
                {
                    throw new Exception($"El libro de Excel seleccionado '{targetOptions?.WorkbookName ?? targetOptions?.WorkbookKey ?? "(sin nombre)"}' ya no está abierto.");
                }

                Logger.LogInfo($"Using open workbook '{GetComStringProperty(workbook, "Name") ?? targetOptions?.WorkbookName ?? "(unknown)"}'");
            }
            else
            {
                // Crear workbook nuevo: excelApp.Workbooks.Add()
                workbook = workbooks?.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, workbooks, null);
                createdWorkbook = workbook != null;
            }

            SafeReleaseComObject(workbooks);

            if (workbook == null)
            {
                Logger.LogError("Failed to create new workbook via Excel.Workbooks.Add");
                throw new Exception("No se pudo crear workbook");
            }
            Logger.LogDebug("Workbook created");
            List<QueryImportDefinition> queryImports = SqlImportQueryPlanner.BuildImportQueries(query);
            string importNameTimestamp = CreateImportNameTimestamp(DateTime.Now);
            bool appendToExistingSheet = useOpenWorkbook && targetOptions?.CreateNewWorksheet == false;
            int rowCount = 0;

            if (queryImports.Count > 1)
            {
                Logger.LogInfo($"Detected {queryImports.Count} final SELECT statements. Creating one Power Query per final SELECT.");
            }

            if (appendToExistingSheet)
            {
                sharedWorksheet = ResolveTargetWorksheet(workbook, targetOptions, useOpenWorkbook);
                if (sharedWorksheet == null)
                {
                    throw new Exception("No se pudo obtener worksheet");
                }
            }

            for (int queryIndex = 0; queryIndex < queryImports.Count; queryIndex++)
            {
                QueryImportDefinition queryImport = queryImports[queryIndex];
                object? currentWorksheet = sharedWorksheet;

                if (currentWorksheet == null)
                {
                    currentWorksheet = ResolveWorksheetForImport(workbook, targetOptions, useOpenWorkbook, queryImport, queryIndex, queryImports.Count, importNameTimestamp);
                }

                if (currentWorksheet == null)
                {
                    throw new Exception("No se pudo obtener worksheet");
                }

                string worksheetName = GetComStringProperty(currentWorksheet, "Name") ?? $"Sheet{queryIndex + 1}";
                Logger.LogInfo($"Importing result query {queryIndex + 1}/{queryImports.Count} into worksheet '{worksheetName}'.");

                try
                {
                    rowCount += ImportQueryIntoWorksheet(workbook, currentWorksheet, server, database, queryImport, importNameTimestamp, appendToExistingSheet);
                }
                finally
                {
                    if (!ReferenceEquals(currentWorksheet, sharedWorksheet))
                    {
                        SafeReleaseComObject(currentWorksheet);
                    }
                }
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
            SafeReleaseComObject(sharedWorksheet);
            
            return new ImportResult
            {
                Success = true,
                RowCount = rowCount,
                QueryCount = queryImports.Count,
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
            SafeReleaseComObject(sharedWorksheet);

            // Cerrar Excel si hubo error
            try
            {
                if (createdWorkbook && workbook != null)
                {
                    workbook.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, workbook, new object[] { false });
                }

                SafeReleaseComObject(workbook);

                if (!attachedToRunningExcel && excelApp != null)
                {
                    excelApp.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, excelApp, null);
                }

                SafeReleaseComObject(excelApp);
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

    private static int ImportQueryIntoWorksheet(
        object workbook,
        object worksheet,
        string server,
        string database,
        QueryImportDefinition queryDefinition,
        string importNameTimestamp,
        bool appendToExistingSheet)
    {
        object? destination = GetDestinationRange(worksheet, appendToExistingSheet);
        string worksheetName = GetComStringProperty(worksheet, "Name") ?? "Sheet1";
        string importObjectName = BuildTimestampedImportName(queryDefinition.SuggestedName, importNameTimestamp);
        string pqName = GetUniqueWorkbookQueryName(workbook, $"Query_{importObjectName}");
        bool powerQueryAdded = false;
        object? powerQueryTable = null;
        Exception? loadException = null;

        Logger.LogDebug($"Obtained destination range for worksheet '{worksheetName}' and query '{pqName}'");

        try
        {
            object? queries = workbook.GetType().InvokeMember("Queries", BindingFlags.GetProperty, null, workbook, null);
            try
            {
                if (queries == null)
                {
                    throw new Exception("Excel no expone la colección de Power Query (Queries) en este libro.");
                }

                string mEscaped = queryDefinition.Script.Replace("\"", "\"\"");
                string mFormula = $"let\r\n    Source = Sql.Database(\"{server}\", \"{database}\", [Query=\"{mEscaped}\"])\r\nin\r\n    Source";
                Logger.LogDebug($"Attempting to add Power Query '{pqName}' to workbook.Queries");
                queries.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, queries, new object[] { pqName, mFormula });
                Logger.LogInfo($"Power Query '{pqName}' added to workbook.Queries");
                powerQueryAdded = true;
            }
            finally
            {
                SafeReleaseComObject(queries);
            }

            powerQueryTable = LoadPowerQueryIntoWorksheet(worksheet, destination, pqName);
            int rowCount = GetImportedRowCount(powerQueryTable);
            Logger.LogInfo($"Power Query '{pqName}' loaded into worksheet with {rowCount} rows.");
            Logger.LogDebug($"ResultRange rows detected: {rowCount}");
            return rowCount;
        }
        catch (TargetInvocationException tie)
        {
            loadException = tie.InnerException ?? tie;
            Logger.LogWarning($"Power Query failed: {loadException.Message}");
            Logger.LogDebug(loadException.ToString());
            TryPromptForNativeQueryPermission(loadException);
        }
        catch (Exception pqEx)
        {
            loadException = pqEx;
            Logger.LogDebug($"Could not complete Power Query import: {pqEx.Message}");
        }
        finally
        {
            SafeReleaseComObject(powerQueryTable);
            SafeReleaseComObject(destination);
        }

        if (!powerQueryAdded)
        {
            throw CreatePowerQueryLoadException("No se pudo crear la Power Query en Excel.", loadException);
        }

        throw CreatePowerQueryLoadException("La Power Query se creó pero no se pudo cargar como tabla en la hoja seleccionada.", loadException);
    }

    private static object LoadPowerQueryIntoWorksheet(object worksheet, object? destination, string pqName)
    {
        object? targetListObjects = null;
        object? createdListObj = null;
        object? createdQtbl = null;

        try
        {
            targetListObjects = worksheet.GetType().InvokeMember("ListObjects", BindingFlags.GetProperty, null, worksheet, null);
            if (targetListObjects == null)
            {
                throw new Exception("No se pudo acceder a la colección de tablas (ListObjects) de la hoja de Excel.");
            }

            if (destination == null)
            {
                throw new Exception("No se pudo calcular la celda de destino para cargar la Power Query.");
            }

            string mashupConnection = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={pqName};Extended Properties=\"\"";
            Logger.LogDebug($"Loading Power Query '{pqName}' into worksheet using Excel Mashup connection.");
            createdListObj = targetListObjects.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, targetListObjects,
                new object[] { 0, mashupConnection, true, 0, destination });

            if (createdListObj == null)
            {
                throw new Exception("Excel no devolvió la tabla (ListObject) al intentar cargar la Power Query.");
            }

            try
            {
                createdListObj.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, createdListObj, new object[] { $"Table_{pqName}" });
            }
            catch (Exception ex)
            {
                Logger.LogDebug($"Could not set ListObject name for Power Query '{pqName}': {ex.Message}");
            }

            createdQtbl = createdListObj.GetType().InvokeMember("QueryTable", BindingFlags.GetProperty, null, createdListObj, null);
            if (createdQtbl == null)
            {
                throw new Exception("Excel no devolvió el QueryTable asociado a la Power Query cargada.");
            }

            ConfigurePowerQueryTable(createdQtbl, $"SELECT * FROM [{pqName}]");
            createdQtbl.GetType().InvokeMember("Refresh", BindingFlags.InvokeMethod, null, createdQtbl, new object[] { false });
            return createdQtbl;
        }
        catch
        {
            TryDeleteListObject(createdListObj, pqName);
            SafeReleaseComObject(createdQtbl);
            throw;
        }
        finally
        {
            SafeReleaseComObject(createdListObj);
            SafeReleaseComObject(targetListObjects);
        }
    }

    private static object? ResolveWorksheetForImport(
        object workbook,
        ImportTargetOptions? targetOptions,
        bool useOpenWorkbook,
        QueryImportDefinition queryDefinition,
        int queryIndex,
        int totalQueries,
        string importNameTimestamp)
    {
        string worksheetBaseName = BuildTimestampedImportName(queryDefinition.SuggestedName, importNameTimestamp);

        if (useOpenWorkbook)
        {
            return AddWorksheet(workbook, GetUniqueWorksheetName(workbook, worksheetBaseName));
        }

        if (queryIndex == 0)
        {
            object? firstWorksheet = GetFirstWorksheet(workbook);
            if (firstWorksheet != null && totalQueries > 1)
            {
                TryRenameWorksheet(firstWorksheet, GetUniqueWorksheetName(workbook, worksheetBaseName));
            }

            return firstWorksheet;
        }

        return AddWorksheet(workbook, GetUniqueWorksheetName(workbook, worksheetBaseName));
    }

    private static object? GetFirstWorksheet(object workbook)
    {
        object? sheets = null;
        try
        {
            sheets = workbook.GetType().InvokeMember("Sheets", BindingFlags.GetProperty, null, workbook, null);
            return sheets?.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, sheets, new object[] { 1 });
        }
        finally
        {
            SafeReleaseComObject(sheets);
        }
    }

    private static void TryRenameWorksheet(object worksheet, string worksheetName)
    {
        try
        {
            worksheet.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, worksheet, new object[] { worksheetName });
        }
        catch (Exception ex)
        {
            Logger.LogDebug($"Could not set worksheet name '{worksheetName}': {ex.Message}");
        }
    }

    private static void ConfigurePowerQueryTable(object queryTable, string commandText)
    {
        try { queryTable.GetType().InvokeMember("CommandType", BindingFlags.SetProperty, null, queryTable, new object[] { 2 }); } catch { }
        try { queryTable.GetType().InvokeMember("CommandText", BindingFlags.SetProperty, null, queryTable, new object[] { new string[] { commandText } }); } catch (Exception ex) { Logger.LogDebug($"Could not set CommandText for query table load: {ex.Message}"); }

        try { queryTable.GetType().InvokeMember("RowNumbers", BindingFlags.SetProperty, null, queryTable, new object[] { false }); } catch { }
        try { queryTable.GetType().InvokeMember("FillAdjacentFormulas", BindingFlags.SetProperty, null, queryTable, new object[] { false }); } catch { }
        try { queryTable.GetType().InvokeMember("RefreshOnFileOpen", BindingFlags.SetProperty, null, queryTable, new object[] { false }); } catch { }
        try { queryTable.GetType().InvokeMember("BackgroundQuery", BindingFlags.SetProperty, null, queryTable, new object[] { false }); } catch { }
        try { queryTable.GetType().InvokeMember("PreserveFormatting", BindingFlags.SetProperty, null, queryTable, new object[] { true }); } catch { }
        try { queryTable.GetType().InvokeMember("RefreshStyle", BindingFlags.SetProperty, null, queryTable, new object[] { 0 }); } catch { }
        try { queryTable.GetType().InvokeMember("SavePassword", BindingFlags.SetProperty, null, queryTable, new object[] { false }); } catch { }
        try { queryTable.GetType().InvokeMember("SaveData", BindingFlags.SetProperty, null, queryTable, new object[] { true }); } catch { }
        try { queryTable.GetType().InvokeMember("AdjustColumnWidth", BindingFlags.SetProperty, null, queryTable, new object[] { true }); } catch { }
    }

    private static void TryDeleteListObject(object? listObject, string pqName)
    {
        try
        {
            if (listObject != null)
            {
                listObject.GetType().InvokeMember("Delete", BindingFlags.InvokeMethod, null, listObject, null);
                Logger.LogDebug($"Deleted empty Power Query table shell for '{pqName}'.");
            }
        }
        catch (Exception ex)
        {
            Logger.LogDebug($"Could not delete empty Power Query table shell: {ex.Message}");
        }
    }

    private static void TryPromptForNativeQueryPermission(Exception exception)
    {
        try
        {
            string message = (exception.Message ?? string.Empty).ToLowerInvariant();
            bool nativeQueryBlocked =
                (message.Contains("evaluate") && message.Contains("native")) ||
                message.Contains("native database") ||
                message.Contains("evaluatenativequeryunpermitted") ||
                message.Contains("permission is required to run this native");

            if (nativeQueryBlocked)
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
        catch
        {
        }
    }

    private static Exception CreatePowerQueryLoadException(string message, Exception? innerException)
    {
        return innerException == null
            ? new Exception(message)
            : new Exception($"{message} Detalle: {innerException.Message}", innerException);
    }

    private static object? ResolveTargetWorksheet(object workbook, ImportTargetOptions? targetOptions, bool useOpenWorkbook)
    {
        if (!useOpenWorkbook)
        {
            object? sheets = null;
            try
            {
                sheets = workbook.GetType().InvokeMember("Sheets", BindingFlags.GetProperty, null, workbook, null);
                return sheets?.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, sheets, new object[] { 1 });
            }
            finally
            {
                SafeReleaseComObject(sheets);
            }
        }

        if (targetOptions == null)
        {
            throw new Exception("No se recibieron opciones de destino para el libro abierto.");
        }

        if (targetOptions.CreateNewWorksheet)
        {
            string importNameTimestamp = CreateImportNameTimestamp(DateTime.Now);
            return AddWorksheet(workbook, GetUniqueWorksheetName(workbook, BuildTimestampedImportName("Import", importNameTimestamp)));
        }

        object? worksheet = FindWorksheetByName(workbook, targetOptions.WorksheetName);
        if (worksheet == null)
        {
            throw new Exception($"La hoja '{targetOptions.WorksheetName ?? "(sin nombre)"}' ya no existe en el libro seleccionado.");
        }

        return worksheet;
    }

    private static object? GetDestinationRange(object worksheet, bool appendToExistingSheet)
    {
        if (!appendToExistingSheet)
        {
            return worksheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, worksheet, new object[] { "A1" });
        }

        try
        {
            object? usedRange = worksheet.GetType().InvokeMember("UsedRange", BindingFlags.GetProperty, null, worksheet, null);
            if (usedRange == null)
            {
                return worksheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, worksheet, new object[] { "A1" });
            }

            object? usedValue = usedRange.GetType().InvokeMember("Value2", BindingFlags.GetProperty, null, usedRange, null);
            int usedRow = GetComIntProperty(usedRange, "Row", 1);
            int usedRowCount = 0;
            int usedColumnCount = 0;

            object? rows = null;
            object? columns = null;

            try
            {
                rows = usedRange.GetType().InvokeMember("Rows", BindingFlags.GetProperty, null, usedRange, null);
                columns = usedRange.GetType().InvokeMember("Columns", BindingFlags.GetProperty, null, usedRange, null);
                usedRowCount = GetComCount(rows);
                usedColumnCount = GetComCount(columns);
            }
            finally
            {
                SafeReleaseComObject(rows);
                SafeReleaseComObject(columns);
                SafeReleaseComObject(usedRange);
            }

            bool sheetLooksEmpty = usedValue == null;
            if (!sheetLooksEmpty && usedValue is string textValue)
            {
                sheetLooksEmpty = string.IsNullOrWhiteSpace(textValue) && usedRowCount <= 1 && usedColumnCount <= 1;
            }

            if (sheetLooksEmpty)
            {
                return worksheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, worksheet, new object[] { "A1" });
            }

            int nextRow = Math.Max(1, usedRow + Math.Max(usedRowCount, 1) + 1);
            return worksheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, worksheet, new object[] { $"A{nextRow}" });
        }
        catch (Exception ex)
        {
            Logger.LogDebug($"Could not calculate append destination. Falling back to A1: {ex.Message}");
            return worksheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, worksheet, new object[] { "A1" });
        }
    }

    private static object? TryGetRunningExcelApplication()
    {
        Type? excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType == null)
        {
            return null;
        }

        try
        {
            Guid clsid = excelType.GUID;
            GetActiveObject(ref clsid, IntPtr.Zero, out object runningExcel);
            return runningExcel;
        }
        catch (COMException ex)
        {
            Logger.LogDebug($"Excel is not currently running: {ex.Message}");
            return null;
        }
    }

    private static object? FindWorkbook(object? workbooks, ImportTargetOptions? targetOptions)
    {
        if (workbooks == null || targetOptions == null)
        {
            return null;
        }

        int workbookCount = GetComCount(workbooks);
        for (int index = 1; index <= workbookCount; index++)
        {
            object? workbook = null;
            try
            {
                workbook = workbooks.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, workbooks, new object[] { index });
                if (workbook == null)
                {
                    continue;
                }

                string workbookName = GetComStringProperty(workbook, "Name") ?? string.Empty;
                string workbookFullName = GetComStringProperty(workbook, "FullName") ?? string.Empty;

                if ((!string.IsNullOrWhiteSpace(targetOptions.WorkbookKey) && string.Equals(workbookFullName, targetOptions.WorkbookKey, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrWhiteSpace(targetOptions.WorkbookKey) && string.Equals(workbookName, targetOptions.WorkbookKey, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrWhiteSpace(targetOptions.WorkbookName) && string.Equals(workbookName, targetOptions.WorkbookName, StringComparison.OrdinalIgnoreCase)))
                {
                    return workbook;
                }
            }
            catch (Exception ex)
            {
                Logger.LogDebug($"Could not evaluate workbook candidate at index {index}: {ex.Message}");
            }

            SafeReleaseComObject(workbook);
        }

        return null;
    }

    private static object AddWorksheet(object workbook, string baseSheetName)
    {
        object? worksheets = null;

        try
        {
            worksheets = workbook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, workbook, null);
            object? newWorksheet = worksheets?.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, worksheets, null);
            if (newWorksheet == null)
            {
                throw new Exception("Excel no devolvió una nueva hoja al intentar crearla.");
            }

            string uniqueSheetName = GetUniqueWorksheetName(workbook, baseSheetName);
            try
            {
                newWorksheet.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, newWorksheet, new object[] { uniqueSheetName });
            }
            catch (Exception ex)
            {
                Logger.LogDebug($"Could not set worksheet name '{uniqueSheetName}': {ex.Message}");
            }

            return newWorksheet;
        }
        finally
        {
            SafeReleaseComObject(worksheets);
        }
    }

    private static object? FindWorksheetByName(object workbook, string? worksheetName)
    {
        if (string.IsNullOrWhiteSpace(worksheetName))
        {
            return null;
        }

        object? worksheets = null;

        try
        {
            worksheets = workbook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, workbook, null);
            if (worksheets == null)
            {
                return null;
            }

            int worksheetCount = GetComCount(worksheets);

            for (int index = 1; index <= worksheetCount; index++)
            {
                object? worksheet = null;
                try
                {
                    worksheet = worksheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, worksheets, new object[] { index });
                    string? currentName = worksheet == null ? null : GetComStringProperty(worksheet, "Name");
                    if (string.Equals(currentName, worksheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        return worksheet;
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogDebug($"Could not inspect worksheet '{worksheetName}' at index {index}: {ex.Message}");
                }

                SafeReleaseComObject(worksheet);
            }
        }
        finally
        {
            SafeReleaseComObject(worksheets);
        }

        return null;
    }

    private static List<string> GetWorksheetNames(object workbook)
    {
        List<string> worksheetNames = new();
        object? worksheets = null;

        try
        {
            worksheets = workbook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, workbook, null);
            if (worksheets == null)
            {
                return worksheetNames;
            }

            int worksheetCount = GetComCount(worksheets);

            for (int index = 1; index <= worksheetCount; index++)
            {
                object? worksheet = null;
                try
                {
                    worksheet = worksheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, worksheets, new object[] { index });
                    string? sheetName = worksheet == null ? null : GetComStringProperty(worksheet, "Name");
                    if (!string.IsNullOrWhiteSpace(sheetName))
                    {
                        worksheetNames.Add(sheetName);
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogDebug($"Could not read worksheet name at index {index}: {ex.Message}");
                }
                finally
                {
                    SafeReleaseComObject(worksheet);
                }
            }
        }
        finally
        {
            SafeReleaseComObject(worksheets);
        }

        return worksheetNames;
    }

    private static string GetUniqueWorkbookQueryName(object workbook, string baseName)
    {
        object? queries = null;
        HashSet<string> existingNames = new(StringComparer.OrdinalIgnoreCase);

        try
        {
            queries = workbook.GetType().InvokeMember("Queries", BindingFlags.GetProperty, null, workbook, null);
            if (queries != null)
            {
                int queryCount = GetComCount(queries);
                for (int index = 1; index <= queryCount; index++)
                {
                    object? query = null;
                    try
                    {
                        query = queries.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, queries, new object[] { index });
                        string? queryName = query == null ? null : GetComStringProperty(query, "Name");
                        if (!string.IsNullOrWhiteSpace(queryName))
                        {
                            existingNames.Add(queryName);
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.LogDebug($"Could not inspect existing workbook query at index {index}: {ex.Message}");
                    }
                    finally
                    {
                        SafeReleaseComObject(query);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Logger.LogDebug($"Could not enumerate workbook queries for uniqueness check: {ex.Message}");
        }
        finally
        {
            SafeReleaseComObject(queries);
        }

        return EnsureUniqueName(baseName, existingNames, 255);
    }

    internal static string CreateImportNameTimestamp(DateTime importTime)
    {
        return importTime.ToString("yyyyMMdd_HHmmss");
    }

    internal static string BuildTimestampedImportName(string? baseName, string importNameTimestamp)
    {
        string normalizedBaseName = string.IsNullOrWhiteSpace(baseName) ? "Import" : baseName.Trim();
        return string.IsNullOrWhiteSpace(importNameTimestamp)
            ? normalizedBaseName
            : $"{normalizedBaseName}_{importNameTimestamp}";
    }

    private static string GetUniqueWorksheetName(object workbook, string baseName)
    {
        HashSet<string> existingNames = new(StringComparer.OrdinalIgnoreCase);
        foreach (string worksheetName in GetWorksheetNames(workbook))
        {
            existingNames.Add(worksheetName);
        }

        string sanitizedBaseName = SanitizeWorksheetName(baseName);
        return EnsureUniqueName(sanitizedBaseName, existingNames, 31);
    }

    private static string EnsureUniqueName(string baseName, HashSet<string> existingNames, int maxLength)
    {
        string normalizedBaseName = string.IsNullOrWhiteSpace(baseName) ? "Import" : baseName.Trim();
        if (normalizedBaseName.Length > maxLength)
        {
            normalizedBaseName = normalizedBaseName[..maxLength];
        }

        string candidate = normalizedBaseName;
        int suffix = 1;

        while (existingNames.Contains(candidate))
        {
            string suffixText = $"_{suffix}";
            int baseLength = Math.Max(1, maxLength - suffixText.Length);
            string trimmedBase = normalizedBaseName.Length > baseLength ? normalizedBaseName[..baseLength] : normalizedBaseName;
            candidate = trimmedBase + suffixText;
            suffix++;
        }

        return candidate;
    }

    private static string SanitizeWorksheetName(string sheetName)
    {
        char[] invalidChars = ['[', ']', ':', '*', '?', '/', '\\'];
        string sanitized = sheetName;
        foreach (char invalidChar in invalidChars)
        {
            sanitized = sanitized.Replace(invalidChar, '_');
        }

        if (sanitized.Length > 31)
        {
            sanitized = sanitized[..31];
        }

        return string.IsNullOrWhiteSpace(sanitized) ? "Import" : sanitized;
    }

    private static string? GetComStringProperty(object target, string propertyName)
    {
        try
        {
            object? value = target.GetType().InvokeMember(propertyName, BindingFlags.GetProperty, null, target, null);
            return value?.ToString();
        }
        catch
        {
            return null;
        }
    }

    private static int GetComIntProperty(object target, string propertyName, int defaultValue = 0)
    {
        try
        {
            object? value = target.GetType().InvokeMember(propertyName, BindingFlags.GetProperty, null, target, null);
            return value is int intValue ? intValue : defaultValue;
        }
        catch
        {
            return defaultValue;
        }
    }

    private static int GetComCount(object? comCollection)
    {
        if (comCollection == null)
        {
            return 0;
        }

        try
        {
            object? count = comCollection.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, comCollection, null);
            return count is int intCount ? intCount : 0;
        }
        catch
        {
            return 0;
        }
    }

    private static void SafeReleaseComObject(object? comObject)
    {
        try
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
        catch
        {
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

    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

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
