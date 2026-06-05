using System.ComponentModel;
using Outlook.COM;

namespace Outlook.ReminderApp;

internal sealed class SyncConfigForm : Form
{
    private readonly BindingList<SyncRule> _rules;
    private readonly DataGridView _grid;
    private readonly SyncScheduler _scheduler;
    private readonly Label _lastRunLabel;
    private readonly Label _lastErrorLabel;
    private readonly ListBox _logList;

    public IReadOnlyList<SyncRule> Rules => _rules.ToList();

    public SyncConfigForm(IEnumerable<SyncRule> rules, SyncScheduler scheduler)
    {
        _scheduler = scheduler;
        Text = "Calendar Sync Settings";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        Width = 900;
        Height = 540;

        _rules = new BindingList<SyncRule>(rules.Select(r => r.Clone()).ToList());

        _grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = false,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            ReadOnly = true,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false
        };

        _grid.Columns.Add(new DataGridViewCheckBoxColumn
        {
            HeaderText = "On",
            DataPropertyName = nameof(SyncRule.Enabled),
            Width = 40
        });
        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Source",
            DataPropertyName = nameof(SyncRule.SourceAccount),
            Width = 180
        });
        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Target",
            DataPropertyName = nameof(SyncRule.TargetAccount),
            Width = 180
        });
        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Mode",
            DataPropertyName = nameof(SyncRule.Mode),
            Width = 70
        });
        _grid.Columns.Add(new DataGridViewCheckBoxColumn
        {
            HeaderText = "Outside Hours",
            DataPropertyName = nameof(SyncRule.OutsideWorkHoursOnly),
            Width = 110
        });
        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Start",
            DataPropertyName = nameof(SyncRule.WorkDayStartHour),
            Width = 55
        });
        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "End",
            DataPropertyName = nameof(SyncRule.WorkDayEndHour),
            Width = 55
        });

        _grid.DataSource = _rules;
        _grid.CellDoubleClick += (_, _) => EditSelected();

        var addBtn = new Button { Text = "Add", Width = 80 };
        var editBtn = new Button { Text = "Edit", Width = 80 };
        var removeBtn = new Button { Text = "Remove", Width = 80 };
        var runBtn = new Button { Text = "Run now", Width = 90 };
        var saveBtn = new Button { Text = "Save", Width = 80 };
        var cancelBtn = new Button { Text = "Cancel", Width = 80 };

        addBtn.Click += (_, _) => AddRule();
        editBtn.Click += (_, _) => EditSelected();
        removeBtn.Click += (_, _) => RemoveSelected();
        runBtn.Click += (_, _) => RunSyncNow();
        saveBtn.Click += (_, _) => { DialogResult = DialogResult.OK; Close(); };
        cancelBtn.Click += (_, _) => { DialogResult = DialogResult.Cancel; Close(); };

        var buttons = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 44,
            FlowDirection = FlowDirection.RightToLeft,
            Padding = new Padding(8)
        };

        buttons.Controls.Add(cancelBtn);
        buttons.Controls.Add(saveBtn);
        buttons.Controls.Add(runBtn);
        buttons.Controls.Add(removeBtn);
        buttons.Controls.Add(editBtn);
        buttons.Controls.Add(addBtn);

        _lastRunLabel = new Label
        {
            AutoSize = true,
            Text = "Last run: never",
            Font = new Font("Segoe UI", 9, FontStyle.Regular)
        };

        _lastErrorLabel = new Label
        {
            AutoSize = true,
            Text = "Last error: none",
            Font = new Font("Segoe UI", 9, FontStyle.Regular)
        };

        _logList = new ListBox
        {
            Dock = DockStyle.Fill,
            IntegralHeight = false,
            Font = new Font("Segoe UI", 8.5f, FontStyle.Regular)
        };

        var logHeader = new Label
        {
            AutoSize = true,
            Text = "Recent sync log",
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };

        var logPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 170,
            ColumnCount = 1,
            RowCount = 4,
            Padding = new Padding(8)
        };
        logPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        logPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        logPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        logPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        logPanel.Controls.Add(_lastRunLabel, 0, 0);
        logPanel.Controls.Add(_lastErrorLabel, 0, 1);
        logPanel.Controls.Add(logHeader, 0, 2);
        logPanel.Controls.Add(_logList, 0, 3);

        Controls.Add(_grid);
        Controls.Add(logPanel);
        Controls.Add(buttons);

        _scheduler.LogsUpdated += OnLogsUpdated;
        UpdateLogView();
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        _scheduler.LogsUpdated -= OnLogsUpdated;
        base.OnFormClosed(e);
    }

    private void OnLogsUpdated(object? sender, EventArgs e)
    {
        if (!IsHandleCreated) return;
        BeginInvoke(UpdateLogView);
    }

    private void UpdateLogView()
    {
        var status = _scheduler.GetStatus();
        _lastRunLabel.Text = status.LastRunAt.HasValue
            ? $"Last run: {status.LastRunAt:yyyy-MM-dd HH:mm:ss}"
            : "Last run: never";

        _lastErrorLabel.Text = status.LastErrorAt.HasValue
            ? $"Last error: {status.LastErrorAt:yyyy-MM-dd HH:mm:ss} - {status.LastError}"
            : "Last error: none";

        var logs = _scheduler.GetRecentLogs();
        _logList.BeginUpdate();
        _logList.Items.Clear();
        foreach (var entry in logs)
        {
            _logList.Items.Add(entry);
        }
        _logList.EndUpdate();
    }

    private void AddRule()
    {
        using var editor = new SyncRuleEditorForm(new SyncRule());
        if (editor.ShowDialog(this) == DialogResult.OK && editor.Rule is not null)
        {
            _rules.Add(editor.Rule);
        }
    }

    private void EditSelected()
    {
        var rule = GetSelectedRule();
        if (rule is null) return;

        using var editor = new SyncRuleEditorForm(rule);
        if (editor.ShowDialog(this) == DialogResult.OK && editor.Rule is not null)
        {
            var index = _rules.IndexOf(rule);
            if (index >= 0)
            {
                _rules[index] = editor.Rule;
            }
        }
    }

    private void RemoveSelected()
    {
        var rule = GetSelectedRule();
        if (rule is null) return;
        _rules.Remove(rule);
    }

    private void RunSyncNow()
    {
        _scheduler.RunAllNow();
        UpdateLogView();
    }

    private SyncRule? GetSelectedRule()
    {
        if (_grid.CurrentRow?.DataBoundItem is SyncRule rule)
        {
            return rule;
        }
        return null;
    }
}
