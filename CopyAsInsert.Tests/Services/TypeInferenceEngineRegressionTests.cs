using CopyAsInsert.Models;
using CopyAsInsert.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace CopyAsInsert.Tests.Services;

/// <summary>
/// Regression tests for real-world data issues
/// </summary>
[TestClass]
public class TypeInferenceEngineRegressionTests
{
    [TestMethod]
    public void TestEuropeanDecimalFormat()
    {
        // Arrange - European format with comma as decimal separator
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo> { new() { ColumnName = "art_PorcRen", SqlType = "INT" } },
            DataRows = new List<string[]>
            {
                new[] { "0,00" },
                new[] { "0,00" },
                new[] { "0,00" },
                new[] { "100,00" },
                new[] { "100,00" },
                new[] { "100,00" },
                new[] { "100,00" },
                new[] { "100,00" },
                new[] { "100,00" },
                new[] { "100,00" }
            }
        };

        // Act
        TypeInferenceEngine.InferColumnTypes(schema);

        // Assert
        Assert.AreEqual("FLOAT", schema.Columns[0].SqlType,
            "Comma-separated decimals (0,00 and 100,00) should be detected as FLOAT, not BIT");
        Assert.IsTrue(schema.Columns[0].ConfidencePercent >= 85);
    }

    [TestMethod]
    public void TestSmallDecimalValues()
    {
        // Arrange - small decimal values with comma separator
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo> { new() { ColumnName = "art_CosteAct", SqlType = "INT" } },
            DataRows = new List<string[]>
            {
                new[] { "0,0673" },
                new[] { "4,5426" },
                new[] { "0,0239" },
                new[] { "0,00" },
                new[] { "0,00" },
                new[] { "0,00" },
                new[] { "0,00" },
                new[] { "0,00" },
                new[] { "0,00" },
                new[] { "0,00" }
            }
        };

        // Act
        TypeInferenceEngine.InferColumnTypes(schema);

        // Assert
        Assert.AreEqual("FLOAT", schema.Columns[0].SqlType,
            "Small decimals like 0,0673 and 4,5426 should be detected as FLOAT");
    }

    [TestMethod]
    public void TestMixedBitAndDecimal()
    {
        // Arrange - mix of 0/1 values and decimal values
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo> { new() { ColumnName = "MixedColumn", SqlType = "INT" } },
            DataRows = new List<string[]>
            {
                new[] { "1" },
                new[] { "0" },
                new[] { "1" },
                new[] { "0" },
                new[] { "1,5" },  // This should prevent BIT detection
                new[] { "0" },
                new[] { "1" },
                new[] { "0" },
                new[] { "1" },
                new[] { "0" }
            }
        };

        // Act
        TypeInferenceEngine.InferColumnTypes(schema);

        // Assert
        Assert.AreNotEqual("BIT", schema.Columns[0].SqlType,
            "Presence of decimal value (1,5) should prevent BIT detection even with 0/1 values");
    }

    [TestMethod]
    public void TestNullAndEmptyStringHandling()
    {
        // Arrange - NULL values marked as empty strings or "NULL"
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo> { new() { ColumnName = "OptionalColumn", SqlType = "INT" } },
            DataRows = new List<string[]>
            {
                new[] { "100,00" },
                new[] { "" },
                new[] { "200,50" },
                new[] { "NULL" },
                new[] { "300,75" },
                new[] { "" },
                new[] { "50,25" },
                new[] { "75,00" },
                new[] { "" },
                new[] { "150,00" }
            }
        };

        // Act
        TypeInferenceEngine.InferColumnTypes(schema);

        // Assert
        Assert.AreEqual("FLOAT", schema.Columns[0].SqlType);
        Assert.IsTrue(schema.Columns[0].AllowNull, "Column with NULL/empty values should allow NULL");
    }

    [TestMethod]
    public void TestIntegerWithCommaDecimalSeparator()
    {
        // Arrange - values that look like integers but with comma format
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo> { new() { ColumnName = "IntColumn", SqlType = "INT" } },
            DataRows = new List<string[]>
            {
                new[] { "1" },
                new[] { "2" },
                new[] { "3" },
                new[] { "100002" },
                new[] { "100006" },
                new[] { "1" },
                new[] { "1" },
                new[] { "100002" },
                new[] { "1" },
                new[] { "1" }
            }
        };

        // Act
        TypeInferenceEngine.InferColumnTypes(schema);

        // Assert
        Assert.AreEqual("INT", schema.Columns[0].SqlType);
    }

    [TestMethod]
    public void TestDatetimeWithNullValues()
    {
        // Arrange - datetime values with some NULL/empty values
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo> { new() { ColumnName = "art_FecModArt", SqlType = "INT" } },
            DataRows = new List<string[]>
            {
                new[] { "2024-01-16 09:46:06.870" },
                new[] { "1899-12-30 00:00:00.000" },
                new[] { "2025-08-27 09:34:08.863" },
                new[] { "2024-01-16 09:45:34.663" },
                new[] { "" },
                new[] { "" },
                new[] { "" },
                new[] { "" },
                new[] { "" },
                new[] { "" }
            }
        };

        // Act
        TypeInferenceEngine.InferColumnTypes(schema);

        // Assert
        Assert.AreEqual("DATETIME2", schema.Columns[0].SqlType,
            "Should detect DATETIME2 despite many NULL values (80% threshold)");
        Assert.IsTrue(schema.Columns[0].AllowNull);
    }

    [TestMethod]
    public void TestStringWithSpaces()
    {
        // Arrange - Text values with leading/trailing spaces
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo> { new() { ColumnName = "art_TipoCoste", SqlType = "INT" } },
            DataRows = new List<string[]>
            {
                new[] { "A" },
                new[] { "A" },
                new[] { "A" },
                new[] { " " },  // Just spaces
                new[] { "A" },
                new[] { "A" },
                new[] { "A" },
                new[] { "A" },
                new[] { "A" },
                new[] { "A" }
            }
        };

        // Act
        TypeInferenceEngine.InferColumnTypes(schema);

        // Assert
        Assert.AreEqual("NVARCHAR", schema.Columns[0].SqlType);
        Assert.IsTrue(schema.Columns[0].AllowNull, "Whitespace-only values should mark AllowNull=true");
    }

    [TestMethod]
    public void TestMaxLengthCalculation()
    {
        // Arrange - Variable length strings
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo> { new() { ColumnName = "ref_Articulo", SqlType = "INT" } },
            DataRows = new List<string[]>
            {
                new[] { "-----" },
                new[] { "." },
                new[] { "......" },
                new[] { "+++++" },
                new[] { "0040271712" },
                new[] { "0040276812" },
                new[] { "0040277412" },
                new[] { "004048" },
                new[] { "0042480012" },
                new[] { "0043180012" }
            }
        };

        // Act
        TypeInferenceEngine.InferColumnTypes(schema);

        // Assert
        Assert.AreEqual("NVARCHAR", schema.Columns[0].SqlType);
        Assert.IsTrue(schema.Columns[0].MaxLength >= 10, 
            $"MaxLength should be at least 10 (longest string '0040271712'), got {schema.Columns[0].MaxLength}");
    }

    [TestMethod]
    public void TestCaseSensitiveNullKeyword()
    {
        // Arrange - NULL keyword in different cases
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo> { new() { ColumnName = "col_Color", SqlType = "INT" } },
            DataRows = new List<string[]>
            {
                new[] { "NULL" },
                new[] { "0" },
                new[] { "null" },
                new[] { "0" },
                new[] { "NULL" },
                new[] { "1" },
                new[] { "0" },
                new[] { "1" },
                new[] { "0" },
                new[] { "1" }
            }
        };

        // Act
        TypeInferenceEngine.InferColumnTypes(schema);

        // Assert
        Assert.IsTrue(schema.Columns[0].AllowNull, "NULL keyword (any case) should be recognized");
    }

    [TestMethod]
    public void TestBitWithOnOffValues()
    {
        // Arrange - Boolean with on/off values
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo> { new() { ColumnName = "IsActive", SqlType = "INT" } },
            DataRows = new List<string[]>
            {
                new[] { "on" },
                new[] { "off" },
                new[] { "on" },
                new[] { "on" },
                new[] { "off" },
                new[] { "on" },
                new[] { "off" },
                new[] { "on" },
                new[] { "on" },
                new[] { "off" }
            }
        };

        // Act
        TypeInferenceEngine.InferColumnTypes(schema);

        // Assert
        Assert.AreEqual("BIT", schema.Columns[0].SqlType,
            "on/off values should be detected as BIT when all values match boolean patterns");
    }
}
