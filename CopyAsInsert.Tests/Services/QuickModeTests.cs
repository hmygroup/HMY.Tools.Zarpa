using CopyAsInsert.Models;
using CopyAsInsert.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace CopyAsInsert.Tests;

/// <summary>
/// Unit tests for Quick Mode SQL generation
/// </summary>
[TestClass]
public class QuickModeTests
{
    [TestMethod]
    public void TestQuickMode_AllColumnsNVarchar100()
    {
        // Arrange: Create a schema with various types
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo>
            {
                new() { ColumnName = "ID", SqlType = "INT", MaxLength = null },
                new() { ColumnName = "Name", SqlType = "NVARCHAR", MaxLength = 50 },
                new() { ColumnName = "Price", SqlType = "FLOAT", MaxLength = null },
                new() { ColumnName = "Created", SqlType = "DATETIME2", MaxLength = null }
            },
            DataRows = new List<string[]>
            {
                new[] { "1", "Product A", "19.99", "2024-01-01" },
                new[] { "2", "Product B", "29.99", "2024-01-02" }
            }
        };

        // Act: Override all columns to NVARCHAR(100) as quick mode does
        foreach (var column in schema.Columns)
        {
            column.SqlType = "NVARCHAR";
            column.MaxLength = 100;
        }

        // Assert: All columns should be NVARCHAR(100)
        foreach (var column in schema.Columns)
        {
            Assert.AreEqual("NVARCHAR", column.SqlType);
            Assert.AreEqual(100, column.MaxLength);
        }
    }

    [TestMethod]
    public void TestQuickMode_GeneratesTempTable()
    {
        // Arrange
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo>
            {
                new() { ColumnName = "Col1", SqlType = "NVARCHAR", MaxLength = 100 },
                new() { ColumnName = "Col2", SqlType = "NVARCHAR", MaxLength = 100 }
            },
            DataRows = new List<string[]>
            {
                new[] { "Value1", "Value2" }
            }
        };

        // Act: Generate SQL with quick mode settings
        var result = SqlServerGenerator.GenerateSql(
            schema,
            tableName: "temp",
            schema_name: "dbo",
            isTemporalTable: false,
            isTemporaryTable: true,
            autoAppendTemporalSuffix: false);

        // Assert
        Assert.IsTrue(result.Success);
        Assert.IsTrue(result.GeneratedSql.Contains("CREATE TABLE [#temp]"));
        Assert.IsTrue(result.GeneratedSql.Contains("[Col1] NVARCHAR(100)"));
        Assert.IsTrue(result.GeneratedSql.Contains("[Col2] NVARCHAR(100)"));
        Assert.IsTrue(result.GeneratedSql.Contains("INSERT INTO [#temp]"));
    }

    [TestMethod]
    public void TestQuickMode_NoTemporalColumns()
    {
        // Arrange
        var schema = new DataTableSchema
        {
            Columns = new List<ColumnTypeInfo>
            {
                new() { ColumnName = "Data", SqlType = "NVARCHAR", MaxLength = 100 }
            },
            DataRows = new List<string[]>
            {
                new[] { "Test" }
            }
        };

        // Act: Generate SQL with quick mode settings (no temporal table)
        var result = SqlServerGenerator.GenerateSql(
            schema,
            tableName: "temp",
            schema_name: "dbo",
            isTemporalTable: false,
            isTemporaryTable: true,
            autoAppendTemporalSuffix: false);

        // Assert: Should not have temporal columns
        Assert.IsTrue(result.Success);
        Assert.IsFalse(result.GeneratedSql.Contains("SysStartTime"));
        Assert.IsFalse(result.GeneratedSql.Contains("SysEndTime"));
        Assert.IsFalse(result.GeneratedSql.Contains("SYSTEM_VERSIONING"));
    }
}
