using CopyAsInsert.Models;
using CopyAsInsert.Services;
using System.ComponentModel;

namespace CopyAsInsert.Forms;

/// <summary>
/// UserControl for reviewing and overriding inferred column types
/// Displays a grid with: Column Name | Inferred Type | Confidence | Actual Type (editable dropdown)
/// </summary>
public partial class TypeOverrideControl : UserControl
{
    private DataTableSchema? _schema;
    private DataGridView? _gridColumnTypes;

    public TypeOverrideControl()
    {
        InitializeComponent();
        InitializeControls();
    }

    private void InitializeComponent()
    {
        this.SuspendLayout();
        
        // Set form properties
        this.AutoScaleDimensions = new SizeF(7F, 15F);
        this.AutoScaleMode = AutoScaleMode.Font;
        this.Size = new Size(800, 400);
        
        this.ResumeLayout(false);
    }

    private void InitializeControls()
    {
        // Grid only (label and button handled by parent form)
        _gridColumnTypes = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = false,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            ReadOnly = false,
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            BackgroundColor = SystemColors.Window
        };

        // Define columns
        _gridColumnTypes.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "ColumnName",
            HeaderText = "Column Name",
            ReadOnly = true,
            Width = 150,
            DataPropertyName = "ColumnName"
        });

        _gridColumnTypes.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "InferredType",
            HeaderText = "Inferred Type",
            ReadOnly = true,
            Width = 100,
            DataPropertyName = "SqlType"
        });

        _gridColumnTypes.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "Confidence",
            HeaderText = "Confidence %",
            ReadOnly = true,
            Width = 80,
            DataPropertyName = "ConfidencePercent"
        });

        _gridColumnTypes.Columns.Add(new DataGridViewComboBoxColumn
        {
            Name = "ActualType",
            HeaderText = "Actual Type",
            Width = 120,
            DataPropertyName = "SqlType",
            Items = { "INT", "FLOAT", "DATETIME2", "NVARCHAR" }
        });

        _gridColumnTypes.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "Reason",
            HeaderText = "Inference Reason",
            ReadOnly = true,
            Width = 300,
            DataPropertyName = "InferenceReason"
        });

        _gridColumnTypes.CellFormatting += GridColumnTypes_CellFormatting;
        this.Controls.Add(_gridColumnTypes);
    }

    /// <summary>
    /// Load schema and populate the grid
    /// </summary>
    public void LoadSchema(DataTableSchema schema)
    {
        _schema = schema;
        RefreshGrid();
    }

    /// <summary>
    /// Refresh the grid with current schema data
    /// </summary>
    private void RefreshGrid()
    {
        if (_gridColumnTypes == null || _schema == null)
            return;

        _gridColumnTypes.DataSource = null;  // Clear existing binding
        _gridColumnTypes.DataSource = new BindingSource(_schema.Columns, null);
    }

    /// <summary>
    /// Apply cell formatting: highlight low-confidence columns in yellow
    /// </summary>
    private void GridColumnTypes_CellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (_gridColumnTypes == null || _gridColumnTypes.Rows.Count <= e.RowIndex || e.RowIndex < 0)
            return;

        var row = _gridColumnTypes.Rows[e.RowIndex];
        var column = row.DataBoundItem as ColumnTypeInfo;

        if (column != null && column.ConfidencePercent < 85)
        {
            // Highlight entire row in light yellow for low confidence
            row.DefaultCellStyle.BackColor = Color.LightYellow;
        }
        else if (column != null)
        {
            row.DefaultCellStyle.BackColor = SystemColors.Window;
        }
    }

    /// <summary>
    /// Get the modified schema with user-selected types
    /// </summary>
    public DataTableSchema? GetModifiedSchema()
    {
        if (_gridColumnTypes?.DataSource is BindingSource binding && binding.DataSource is List<ColumnTypeInfo> columns)
        {
            // Update schema with any user-modified types from grid
            foreach (DataGridViewRow row in _gridColumnTypes.Rows)
            {
                if (row.Cells["ActualType"].Value != null)
                {
                    string selectedType = row.Cells["ActualType"].Value.ToString() ?? "NVARCHAR";
                    var column = columns[row.Index];
                    column.SqlType = selectedType;
                }
            }
        }

        return _schema;
    }

    /// <summary>
    /// Set all columns to NVARCHAR type
    /// </summary>
    public void SetAllColumnsToNvarchar()
    {
        if (_gridColumnTypes == null)
            return;

        foreach (DataGridViewRow row in _gridColumnTypes.Rows)
        {
            row.Cells["ActualType"].Value = "NVARCHAR";
        }
    }
}
