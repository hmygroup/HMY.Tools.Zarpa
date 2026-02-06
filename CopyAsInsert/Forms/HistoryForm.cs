using CopyAsInsert.Models;
using CopyAsInsert.Services;

namespace CopyAsInsert.Forms;

/// <summary>
/// History dialog displaying all conversion records in a data grid
/// </summary>
public partial class HistoryForm : Form
{
    private readonly List<ConversionResult> _history;
    private DataGridView _dataGridView = null!;
    private ContextMenuStrip _contextMenu = null!;

    public HistoryForm(List<ConversionResult> history)
    {
        _history = history ?? new List<ConversionResult>();
        InitializeComponent();
        PopulateGrid();
    }

    private void InitializeComponent()
    {
        this.Text = "Conversion History";
        this.Width = 900;
        this.Height = 500;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.MaximizeBox = true;
        this.MinimizeBox = false;
        this.ShowIcon = false;

        // Main layout
        var mainLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Padding = new Padding(10)
        };
        mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));

        // DataGridView
        _dataGridView = new DataGridView
        {
            Name = "dgvHistory",
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            ReadOnly = true,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false,
            BorderStyle = BorderStyle.Fixed3D,
            BackgroundColor = SystemColors.Window,
            RowHeadersVisible = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        };

        // Define columns
        var dtColumn = new DataGridViewTextBoxColumn
        {
            Name = "Timestamp",
            HeaderText = "Timestamp",
            DataPropertyName = "ConversionTime",
            DefaultCellStyle = new DataGridViewCellStyle
            {
                Format = "g", // General format with time
            },
            FillWeight = 20
        };

        var tableColumn = new DataGridViewTextBoxColumn
        {
            Name = "TableName",
            HeaderText = "Table Name",
            DataPropertyName = "TableName",
            FillWeight = 15
        };

        var rowsColumn = new DataGridViewTextBoxColumn
        {
            Name = "RowCount",
            HeaderText = "Rows",
            DataPropertyName = "RowCount",
            FillWeight = 8
        };

        var sqlColumn = new DataGridViewTextBoxColumn
        {
            Name = "SqlPreview",
            HeaderText = "SQL Preview",
            DataPropertyName = "GeneratedSql",
            FillWeight = 57
        };

        _dataGridView.Columns.AddRange(dtColumn, tableColumn, rowsColumn, sqlColumn);

        // Context menu
        _contextMenu = new ContextMenuStrip();
        _contextMenu.Items.Add(new ToolStripMenuItem("Copy SQL", null, (s, e) => CopySql()));
        _contextMenu.Items.Add(new ToolStripMenuItem("Copy Summary", null, (s, e) => CopySummary()));
        _contextMenu.Items.Add(new ToolStripSeparator());
        _contextMenu.Items.Add(new ToolStripMenuItem("Delete", null, (s, e) => DeleteRow()));

        _dataGridView.ContextMenuStrip = _contextMenu;

        // Buttons panel
        var buttonPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(5)
        };

        var closeButton = new Button
        {
            Text = "Close",
            DialogResult = DialogResult.OK,
            Width = 80,
            Height = 30,
            Anchor = AnchorStyles.Right | AnchorStyles.Bottom
        };
        closeButton.Left = buttonPanel.Width - closeButton.Width - 5;
        closeButton.Top = buttonPanel.Height - closeButton.Height - 5;

        buttonPanel.Controls.Add(closeButton);

        mainLayout.Controls.Add(_dataGridView, 0, 0);
        mainLayout.Controls.Add(buttonPanel, 0, 1);

        this.Controls.Add(mainLayout);
        this.AcceptButton = closeButton;
    }

    private void PopulateGrid()
    {
        _dataGridView.DataSource = null;
        _dataGridView.Rows.Clear();

        var displayList = _history.Select(x => new
        {
            x.ConversionTime,
            x.TableName,
            x.RowCount,
            SqlPreview = GetSqlPreview(x.GeneratedSql),
            FullSql = x.GeneratedSql,
            Summary = $"{x.TableName} ({x.RowCount} rows)"
        }).ToList();

        // Populate grid manually to avoid binding issues
        foreach (var item in displayList)
        {
            int rowIndex = _dataGridView.Rows.Add(
                item.ConversionTime.ToString("g"),
                item.TableName,
                item.RowCount.ToString(),
                item.SqlPreview
            );

            var row = _dataGridView.Rows[rowIndex];
            row.Tag = item;
            row.Height = 40;
        }

        // Adjust column widths for better readability
        if (_dataGridView.Columns.Count >= 4)
        {
            _dataGridView.Columns[0].Width = 150;  // Timestamp
            _dataGridView.Columns[1].Width = 120;  // TableName
            _dataGridView.Columns[2].Width = 60;   // RowCount
            _dataGridView.Columns[3].Width = 500;  // SqlPreview
        }
    }

    private string GetSqlPreview(string sql)
    {
        const int maxLength = 100;
        if (string.IsNullOrEmpty(sql))
            return string.Empty;

        return sql.Length > maxLength ? sql.Substring(0, maxLength) + "..." : sql;
    }

    private void CopySql()
    {
        if (_dataGridView.SelectedRows.Count > 0)
        {
            var row = _dataGridView.SelectedRows[0];
            var tag = row.Tag as dynamic;
            if (tag != null)
            {
                string fullSql = tag.FullSql;
                ClipboardInterceptor.SetClipboardText(fullSql);
                MessageBox.Show("SQL copied to clipboard", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Logger.LogInfo("History: SQL copied to clipboard");
            }
        }
    }

    private void CopySummary()
    {
        if (_dataGridView.SelectedRows.Count > 0)
        {
            var row = _dataGridView.SelectedRows[0];
            var tag = row.Tag as dynamic;
            if (tag != null)
            {
                string summary = tag.Summary;
                Clipboard.SetText(summary);
                MessageBox.Show($"Copied: {summary}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Logger.LogInfo("History: Summary copied to clipboard");
            }
        }
    }

    private void DeleteRow()
    {
        if (_dataGridView.SelectedRows.Count > 0)
        {
            var rowIndex = _dataGridView.SelectedRows[0].Index;
            if (rowIndex >= 0 && rowIndex < _history.Count)
            {
                var result = MessageBox.Show(
                    "Are you sure you want to delete this history entry?",
                    "Confirm Delete",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    _history.RemoveAt(rowIndex);
                    HistoryManager.SaveHistory(_history);
                    PopulateGrid();
                    Logger.LogInfo("History: Entry deleted");
                }
            }
        }
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _dataGridView?.Dispose();
            _contextMenu?.Dispose();
        }
        base.Dispose(disposing);
    }
}
