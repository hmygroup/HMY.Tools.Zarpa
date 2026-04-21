using System;
using CopyAsInsert.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CopyAsInsert.Tests.Services;

[TestClass]
public class SqlImportQueryPlannerTests
{
    [TestMethod]
    public void BuildImportQueries_SplitsMultipleTrailingFinalSelects()
    {
        const string sql = """
DROP TABLE IF EXISTS #CAB;
SELECT cab.Id, agg.MaxFecha
INTO #CAB
FROM dbo.Cabecera cab
LEFT JOIN (
    SELECT DetId, MAX(Fecha) AS MaxFecha
    FROM dbo.Detalle
    GROUP BY DetId
) agg
    ON agg.DetId = cab.Id;

DROP TABLE IF EXISTS #LIN;
SELECT Pedido
INTO #LIN
FROM #CAB;

SELECT *
FROM #CAB

SELECT *
FROM #LIN
""";

        var plans = SqlImportQueryPlanner.BuildImportQueries(sql);

        Assert.AreEqual(2, plans.Count);
        StringAssert.Contains(plans[0].Script, "DROP TABLE IF EXISTS #CAB;");
        StringAssert.Contains(plans[0].Script, "SELECT *\nFROM #CAB".Replace("\n", Environment.NewLine));
        Assert.IsFalse(plans[0].Script.Contains("FROM #LIN", StringComparison.OrdinalIgnoreCase));
        Assert.AreEqual("CAB", plans[0].SuggestedName);

        StringAssert.Contains(plans[1].Script, "DROP TABLE IF EXISTS #LIN;");
        StringAssert.Contains(plans[1].Script, "SELECT *\nFROM #LIN".Replace("\n", Environment.NewLine));
        Assert.IsFalse(plans[1].Script.Contains("FROM #CAB" + Environment.NewLine + Environment.NewLine + "SELECT *", StringComparison.OrdinalIgnoreCase));
        Assert.AreEqual("LIN", plans[1].SuggestedName);
    }

    [TestMethod]
    public void BuildImportQueries_LeavesSingleFinalSelectUntouched()
    {
        const string sql = """
DROP TABLE IF EXISTS #CAB;
SELECT Id
INTO #CAB
FROM dbo.Cabecera;

SELECT *
FROM #CAB
""";

        var plans = SqlImportQueryPlanner.BuildImportQueries(sql);

        Assert.AreEqual(1, plans.Count);
        Assert.AreEqual(sql.Trim(), plans[0].Script);
    }

    [TestMethod]
    public void BuildImportQueries_IgnoresCommentedSelectsAtEnd()
    {
        const string sql = """
SELECT *
FROM dbo.RealTable

-- SELECT *
-- FROM dbo.CommentedTable
""";

        var plans = SqlImportQueryPlanner.BuildImportQueries(sql);

        Assert.AreEqual(1, plans.Count);
        Assert.AreEqual(sql.Trim(), plans[0].Script);
    }
}