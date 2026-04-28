using System.Collections.Generic;
using CopyAsInsert.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CopyAsInsert.Tests.Services;

[TestClass]
public class SettingsManagerRecentValuesTests
{
    [TestMethod]
    public void UpdateExcelImportRecentValues_PutsLatestValuesFirst()
    {
        var settings = new SettingsManager.ApplicationSettings
        {
            ExcelImportServerHistory = new List<string> { "srv-old", "srv-older" },
            ExcelImportDatabaseHistory = new List<string> { "db-old", "db-older" }
        };

        SettingsManager.UpdateExcelImportRecentValues(settings, "srv-new", "db-new");

        CollectionAssert.AreEqual(new List<string> { "srv-new", "srv-old", "srv-older" }, settings.ExcelImportServerHistory);
        CollectionAssert.AreEqual(new List<string> { "db-new", "db-old", "db-older" }, settings.ExcelImportDatabaseHistory);
    }

    [TestMethod]
    public void UpdateExcelImportRecentValues_DeduplicatesIgnoringCaseAndWhitespace()
    {
        var settings = new SettingsManager.ApplicationSettings
        {
            ExcelImportServerHistory = new List<string> { " SQL01 ", "sql02" },
            ExcelImportDatabaseHistory = new List<string> { "Ventas", "Compras" }
        };

        SettingsManager.UpdateExcelImportRecentValues(settings, "sql01", " ventas ");

        CollectionAssert.AreEqual(new List<string> { "sql01", "sql02" }, settings.ExcelImportServerHistory);
        CollectionAssert.AreEqual(new List<string> { "ventas", "Compras" }, settings.ExcelImportDatabaseHistory);
    }
}
