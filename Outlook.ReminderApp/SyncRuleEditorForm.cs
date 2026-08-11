using Outlook.COM;

namespace Outlook.ReminderApp;

internal sealed class SyncRuleEditorForm : Form
{
    private readonly ComboBox _sourceBox;
    private readonly ComboBox _targetBox;
    private readonly ComboBox _modeBox;
    private readonly CheckBox _outsideHoursBox;
    private readonly NumericUpDown _startHour;
    private readonly NumericUpDown _endHour;
    private readonly CheckBox _enabledBox;

    public SyncRule? Rule { get; private set; }

    public SyncRuleEditorForm(SyncRule rule)
    {
        Text = "Sync Rule";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        Width = 420;
        Height = 320;

        var working = rule.Clone();

        var sourceLabel = new Label { Text = "Source account", AutoSize = true };
        var targetLabel = new Label { Text = "Target account", AutoSize = true };
        var modeLabel = new Label { Text = "Mode", AutoSize = true };
        var outsideLabel = new Label { Text = "Outside work hours", AutoSize = true };
        var hoursLabel = new Label { Text = "Work hours", AutoSize = true };

        _sourceBox = new ComboBox { Width = 260, DropDownStyle = ComboBoxStyle.DropDown };
        _targetBox = new ComboBox { Width = 260, DropDownStyle = ComboBoxStyle.DropDown };
        LoadAccounts(_sourceBox, working.SourceAccount);
        LoadAccounts(_targetBox, working.TargetAccount);

        _modeBox = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Width = 120
        };
        _modeBox.Items.Add(SyncMode.Block);
        _modeBox.Items.Add(SyncMode.Copy);
        _modeBox.SelectedItem = working.Mode;

        _outsideHoursBox = new CheckBox { Checked = working.OutsideWorkHoursOnly };

        _startHour = new NumericUpDown { Minimum = 0, Maximum = 23, Width = 60, Value = working.WorkDayStartHour };
        _endHour = new NumericUpDown { Minimum = 0, Maximum = 23, Width = 60, Value = working.WorkDayEndHour };

        _enabledBox = new CheckBox { Text = "Enabled", Checked = working.Enabled, AutoSize = true };

        var saveBtn = new Button { Text = "OK", Width = 80 };
        var cancelBtn = new Button { Text = "Cancel", Width = 80 };

        saveBtn.Click += (_, _) =>
        {
            var source = GetComboText(_sourceBox);
            var target = GetComboText(_targetBox);
            if (string.IsNullOrWhiteSpace(source) || string.IsNullOrWhiteSpace(target))
            {
                MessageBox.Show(this, "Source and target are required.", "Sync Rule", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Rule = new SyncRule
            {
                Enabled = _enabledBox.Checked,
                SourceAccount = source,
                TargetAccount = target,
                Mode = (SyncMode)_modeBox.SelectedItem!,
                OutsideWorkHoursOnly = _outsideHoursBox.Checked,
                WorkDayStartHour = (int)_startHour.Value,
                WorkDayEndHour = (int)_endHour.Value
            };

            DialogResult = DialogResult.OK;
            Close();
        };

        cancelBtn.Click += (_, _) =>
        {
            DialogResult = DialogResult.Cancel;
            Close();
        };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 7,
            Padding = new Padding(12),
            AutoSize = true
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));

        layout.Controls.Add(sourceLabel, 0, 0);
        layout.Controls.Add(_sourceBox, 1, 0);
        layout.Controls.Add(targetLabel, 0, 1);
        layout.Controls.Add(_targetBox, 1, 1);
        layout.Controls.Add(modeLabel, 0, 2);
        layout.Controls.Add(_modeBox, 1, 2);
        layout.Controls.Add(outsideLabel, 0, 3);
        layout.Controls.Add(_outsideHoursBox, 1, 3);
        layout.Controls.Add(hoursLabel, 0, 4);

        var hoursPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true };
        hoursPanel.Controls.Add(_startHour);
        hoursPanel.Controls.Add(new Label { Text = "to", AutoSize = true, Padding = new Padding(6, 6, 6, 0) });
        hoursPanel.Controls.Add(_endHour);
        layout.Controls.Add(hoursPanel, 1, 4);

        layout.Controls.Add(_enabledBox, 1, 5);

        var buttons = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.RightToLeft,
            Dock = DockStyle.Bottom,
            Height = 44,
            Padding = new Padding(12)
        };
        buttons.Controls.Add(cancelBtn);
        buttons.Controls.Add(saveBtn);

        Controls.Add(layout);
        Controls.Add(buttons);
    }

    private static void LoadAccounts(ComboBox comboBox, string currentValue)
    {
        comboBox.Items.Clear();
        comboBox.Items.Add("Loading...");
        comboBox.SelectedIndex = 0;

        if (!string.IsNullOrWhiteSpace(currentValue))
        {
            comboBox.Text = currentValue;
        }

        // Doesn't touch COM directly — OutlookCalendarService routes every real call through
        // Outlook.COM's dedicated STA worker (see ComTimeout), so this thread doesn't need to
        // be STA itself.
        Task.Run(() =>
        {
            List<string> names = new();
            try
            {
                var calService = new OutlookCalendarService();
                var accounts = calService.ListAccounts();
                foreach (var account in accounts)
                {
                    if (account.TryGetValue("displayName", out var nameObj) && nameObj is string name && !string.IsNullOrWhiteSpace(name))
                    {
                        names.Add(name);
                    }
                }
            }
            catch
            {
                // Fall back to manual entry if Outlook is unavailable.
            }

            if (comboBox.IsDisposed) return;
            comboBox.BeginInvoke(() =>
            {
                if (comboBox.IsDisposed) return;
                comboBox.Items.Clear();
                foreach (var name in names)
                {
                    comboBox.Items.Add(name);
                }

                if (!string.IsNullOrWhiteSpace(currentValue))
                {
                    comboBox.Text = currentValue;
                }
            });
        });
    }

    private static string GetComboText(ComboBox comboBox)
    {
        return comboBox.Text.Trim();
    }
}
