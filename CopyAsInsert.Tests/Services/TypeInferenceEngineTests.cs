using CopyAsInsert.Models;
using CopyAsInsert.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace CopyAsInsert.Tests;

/// <summary>
/// Unit tests for TypeInferenceEngine - tests all type detection scenarios
/// </summary>
[TestClass]
public class TypeInferenceEngineTests
{
    [TestMethod]
    public void TestIntegerDetection()
    {
        // Arrange
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo> { new() { ColumnName = "ID", SqlType = "INT" } },
            DataRows = new List<string[]>
            {
                new[] { "1" }, new[] { "2" }, new[] { "3" }, new[] { "100" }, new[] { "500" },
                new[] { "1000" }, new[] { "9999" }, new[] { "12345" }, new[] { "99999" }, new[] { "1" }
            }
        };

        // Act
        TypeInferenceEngine.InferColumnTypes(schema);

        // Assert
        Assert.AreEqual("INT", schema.Columns[0].SqlType);
        Assert.IsTrue(schema.Columns[0].ConfidencePercent >= 85);
    }

    [TestMethod]
    public void TestFloatDetection()
    {
        // Arrange
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo> { new() { ColumnName = "Price", SqlType = "INT" } },
            DataRows = new List<string[]>
            {
                new[] { "1.5" }, new[] { "2.99" }, new[] { "10.50" }, new[] { "99.99" }, new[] { "0.99" },
                new[] { "15.75" }, new[] { "200.00" }, new[] { "50.25" }, new[] { "123.45" }, new[] { "5.5" }
            }
        };

        // Act
        TypeInferenceEngine.InferColumnTypes(schema);

        // Assert
        Assert.AreEqual("FLOAT", schema.Columns[0].SqlType);
        Assert.IsTrue(schema.Columns[0].ConfidencePercent >= 85);
    }

    [TestMethod]
    public void TestDateTimeDetection()
    {
        // Arrange
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo> { new() { ColumnName = "CreatedDate", SqlType = "INT" } },
            DataRows = new List<string[]>
            {
                new[] { "2026-02-06" }, new[] { "2026-02-05" }, new[] { "2026-02-04" }, new[] { "2026-02-03" }, new[] { "2026-02-02" },
                new[] { "2026-02-01" }, new[] { "2026-01-31" }, new[] { "2026-01-30" }, new[] { "2026-01-29" }, new[] { "2026-01-28" }
            }
        };

        // Act
        TypeInferenceEngine.InferColumnTypes(schema);

        // Assert
        Assert.AreEqual("DATETIME2", schema.Columns[0].SqlType);
        Assert.IsTrue(schema.Columns[0].ConfidencePercent >= 80);
    }

    [TestMethod]
    public void TestBooleanDetection()
    {
        // Arrange
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo> { new() { ColumnName = "IsActive", SqlType = "INT" } },
            DataRows = new List<string[]>
            {
                new[] { "true" }, new[] { "false" }, new[] { "true" }, new[] { "true" }, new[] { "false" },
                new[] { "true" }, new[] { "false" }, new[] { "true" }, new[] { "true" }, new[] { "false" }
            }
        };

        // Act
        TypeInferenceEngine.InferColumnTypes(schema);

        // Assert
        Assert.AreEqual("BIT", schema.Columns[0].SqlType);
        Assert.IsTrue(schema.Columns[0].ConfidencePercent >= 90);
    }

    [TestMethod]
    public void TestMixedBooleanDetection()
    {
        // Arrange - mix of true/false and yes/no
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo> { new() { ColumnName = "Status", SqlType = "INT" } },
            DataRows = new List<string[]>
            {
                new[] { "yes" }, new[] { "no" }, new[] { "yes" }, new[] { "yes" }, new[] { "no" },
                new[] { "yes" }, new[] { "no" }, new[] { "yes" }, new[] { "yes" }, new[] { "no" }
            }
        };

        // Act
        TypeInferenceEngine.InferColumnTypes(schema);

        // Assert
        Assert.AreEqual("BIT", schema.Columns[0].SqlType);
        Assert.IsTrue(schema.Columns[0].ConfidencePercent >= 90);
    }

    [TestMethod]
    public void TestNVarcharFallback()
    {
        // Arrange - mixed text and numbers
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo> { new() { ColumnName = "Description", SqlType = "INT" } },
            DataRows = new List<string[]>
            {
                new[] { "Product A" }, new[] { "123" }, new[] { "Item 2" }, new[] { "456" }, new[] { "Test" },
                new[] { "789" }, new[] { "Alpha" }, new[] { "Beta" }, new[] { "Gamma" }, new[] { "Delta" }
            }
        };

        // Act
        TypeInferenceEngine.InferColumnTypes(schema);

        // Assert
        Assert.AreEqual("NVARCHAR", schema.Columns[0].SqlType);
    }

    [TestMethod]
    public void TestNullHandling()
    {
        // Arrange - integers with some NULL values
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo> { new() { ColumnName = "Count", SqlType = "INT" } },
            DataRows = new List<string[]>
            {
                new[] { "10" }, new[] { "20" }, new[] { "30" }, new[] { "" }, new[] { "40" },
                new[] { "50" }, new[] { "" }, new[] { "60" }, new[] { "70" }, new[] { "80" }
            }
        };

        // Act
        TypeInferenceEngine.InferColumnTypes(schema);

        // Assert
        Assert.AreEqual("INT", schema.Columns[0].SqlType);
        Assert.IsTrue(schema.Columns[0].AllowNull);
    }

    [TestMethod]
    public void TestLeadingZeroHandling()
    {
        // Arrange - integers with leading zeros (like ZIP codes or IDs with padding)
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo> { new() { ColumnName = "ProductID", SqlType = "INT" } },
            DataRows = new List<string[]>
            {
                new[] { "0001" }, new[] { "0002" }, new[] { "0100" }, new[] { "0500" }, new[] { "1000" },
                new[] { "0050" }, new[] { "0075" }, new[] { "0999" }, new[] { "0005" }, new[] { "0010" }
            }
        };

        // Act
        TypeInferenceEngine.InferColumnTypes(schema);

        // Assert
        Assert.AreEqual("INT", schema.Columns[0].SqlType,
            "Leading zeros should not prevent INT detection if values are valid integers");
    }

    [TestMethod]
    public void TestConfidenceScoring()
    {
        // Arrange - 90% integers, 10% text
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo> { new() { ColumnName = "Value", SqlType = "INT" } },
            DataRows = new List<string[]>
            {
                new[] { "1" }, new[] { "2" }, new[] { "3" }, new[] { "4" }, new[] { "5" },
                new[] { "6" }, new[] { "7" }, new[] { "8" }, new[] { "9" }, new[] { "TEXT" }
            }
        };

        // Act
        TypeInferenceEngine.InferColumnTypes(schema);

        // Assert
        Assert.AreEqual("INT", schema.Columns[0].SqlType);
        Assert.IsTrue(schema.Columns[0].ConfidencePercent == 90,
            $"Expected 90% confidence but got {schema.Columns[0].ConfidencePercent}%");
    }

    [TestMethod]
    public void TestInferenceReason()
    {
        // Arrange
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo> { new() { ColumnName = "Count", SqlType = "INT" } },
            DataRows = new List<string[]>
            {
                new[] { "100" }, new[] { "200" }, new[] { "300" }
            }
        };

        // Act
        TypeInferenceEngine.InferColumnTypes(schema);

        // Assert
        Assert.IsNotNull(schema.Columns[0].InferenceReason);
        Assert.IsTrue(schema.Columns[0].InferenceReason.Length > 0);
        StringAssert.Contains(schema.Columns[0].InferenceReason, "100", 
            "Inference reason should include the confidence percentage");
    }

    [TestMethod]
    public void TestMaxLengthForNVarchar()
    {
        // Arrange
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo> { new() { ColumnName = "Name", SqlType = "INT" } },
            DataRows = new List<string[]>
            {
                new[] { "Apple" }, new[] { "Banana" }, new[] { "Supercalifragilisticexpialidocious" }
            }
        };

        // Act
        TypeInferenceEngine.InferColumnTypes(schema);

        // Assert
        Assert.AreEqual("NVARCHAR", schema.Columns[0].SqlType);
        Assert.IsTrue(schema.Columns[0].MaxLength >= 34, 
            $"Max length should be at least 34 but got {schema.Columns[0].MaxLength}");
    }

    [TestMethod]
    public void TestEmptyColumn()
    {
        // Arrange
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo> { new() { ColumnName = "Empty", SqlType = "INT" } },
            DataRows = new List<string[]>
            {
                new[] { "" }, new[] { "" }, new[] { "" }, new[] { "" }
            }
        };

        // Act
        TypeInferenceEngine.InferColumnTypes(schema);

        // Assert
        Assert.AreEqual("NVARCHAR", schema.Columns[0].SqlType);
        Assert.IsTrue(schema.Columns[0].AllowNull);
    }
}
