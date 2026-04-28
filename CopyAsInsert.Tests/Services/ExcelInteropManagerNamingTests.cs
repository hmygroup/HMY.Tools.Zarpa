using System;
using CopyAsInsert.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CopyAsInsert.Tests.Services;

[TestClass]
public class ExcelInteropManagerNamingTests
{
    [TestMethod]
    public void CreateImportNameTimestamp_IncludesDayAndSeconds()
    {
        var importTime = new DateTime(2026, 4, 28, 15, 30, 45);

        string timestamp = ExcelInteropManager.CreateImportNameTimestamp(importTime);

        Assert.AreEqual("20260428_153045", timestamp);
    }

    [TestMethod]
    public void BuildTimestampedImportName_AppendsTimestampToBaseName()
    {
        string importName = ExcelInteropManager.BuildTimestampedImportName("CAB", "20260428_153045");

        Assert.AreEqual("CAB_20260428_153045", importName);
    }

    [TestMethod]
    public void BuildTimestampedImportName_UsesImportWhenBaseNameMissing()
    {
        string importName = ExcelInteropManager.BuildTimestampedImportName("   ", "20260428_153045");

        Assert.AreEqual("Import_20260428_153045", importName);
    }
}
