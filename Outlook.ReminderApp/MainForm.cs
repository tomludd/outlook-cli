namespace Outlook.ReminderApp;

internal sealed class MainForm : Form
{
    private static readonly IntPtr HwndTopMost = new(-1);
    private const int WsExToolWindow = 0x00000080;
    private const int WsExAppWindow = 0x00040000;
    private const uint SwpNoSize = 0x0001;
    private const uint SwpNoMove = 0x0002;
    private const uint SwpNoActivate = 0x0010;
    private const uint SwpShowWindow = 0x0040;

    private readonly MeetingReminderService _reminderService;
    private readonly System.Windows.Forms.Timer _timer;
    private readonly FlowLayoutPanel _meetingList;
    private readonly NotifyIcon _trayIcon;
    private readonly Icon _trayBellIcon;

    private DateTime _nextRefreshAt = DateTime.MinValue;
    private IReadOnlyList<ReminderMeeting> _currentMeetings = Array.Empty<ReminderMeeting>();
    private ReminderMeeting? _nextMeeting;
    private IReadOnlyList<string> _renderedMeetingIds = Array.Empty<string>();

    public MainForm(MeetingReminderService reminderService)
    {
        _reminderService = reminderService;

        AutoScaleMode = AutoScaleMode.None;
        FormBorderStyle = FormBorderStyle.None;
        StartPosition = FormStartPosition.Manual;
        ShowInTaskbar = false;
        var appIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        if (appIcon is not null) Icon = appIcon;
        TopMost = true;
        BackColor = Color.Magenta;
        TransparencyKey = Color.Magenta;
        ForeColor = Color.WhiteSmoke;
        Width = 550;
        Height = 48;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint, true);
        UpdateStyles();

        // Listen for display configuration changes (monitor plug/unplug, laptop close/open, resolution changes)
        Microsoft.Win32.SystemEvents.DisplaySettingsChanged += OnDisplaySettingsChanged;

        var container = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(8),
            BackColor = Color.Transparent
        };

        _meetingList = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoScroll = false,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            BackColor = Color.Transparent
        };

        container.Controls.Add(_meetingList);
        Controls.Add(container);

        _timer = new System.Windows.Forms.Timer
        {
            Interval = 1000
        };
        _timer.Tick += (_, _) => OnTick();

        var trayMenu = new ContextMenuStrip();
        trayMenu.Items.Add("Exit", null, (_, _) => Application.Exit());

        _trayBellIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath) ?? CreateTrayBellIcon();

        _trayIcon = new NotifyIcon
        {
            Icon = _trayBellIcon,
            Text = "Meeting Reminder",
            Visible = true,
            ContextMenuStrip = trayMenu
        };

        Shown += (_, _) => EnsureTopMostWindow();
        Activated += (_, _) => EnsureTopMostWindow();
        Deactivate += (_, _) => BeginInvoke(EnsureTopMostWindow);

        Load += (_, _) =>
        {
            _timer.Start();
            RefreshMeetings(DateTime.Now);
            RebuildView(DateTime.Now);
            EnsureTopMostWindow();
        };
    }

    protected override CreateParams CreateParams
    {
        get
        {
            var createParams = base.CreateParams;
            createParams.ExStyle |= WsExToolWindow;
            createParams.ExStyle &= ~WsExAppWindow;
            return createParams;
        }
    }

    private void OnTick()
    {
        var now = DateTime.Now;
        var didRefresh = false;

        if (now >= _nextRefreshAt)
        {
            RefreshMeetings(now);
            didRefresh = true;
        }

        _reminderService.TryAutoOpenDueMeetings(_currentMeetings, now);

        if (didRefresh)
        {
            var newIds = GetMeetingsToShow().Select(m => m.Id).ToList();
            if (!newIds.SequenceEqual(_renderedMeetingIds))
            {
                RebuildView(now);
            }
            else
            {
                UpdateRowCountdowns(now);
                UpdateTrayIcon(now);
            }
        }
        else
        {
            UpdateRowCountdowns(now);
            UpdateTrayIcon(now);
        }
    }

    private void RefreshMeetings(DateTime now)
    {
        try
        {
            _currentMeetings = _reminderService.GetVisibleMeetings(now);
            _nextMeeting = _reminderService.GetNextMeeting(now);
        }
        catch
        {
            _currentMeetings = Array.Empty<ReminderMeeting>();
            _nextMeeting = null;
        }

        _nextRefreshAt = now.AddSeconds(15);
    }

    private void RebuildView(DateTime now)
    {
        var meetingsToShow = GetMeetingsToShow();

        _renderedMeetingIds = meetingsToShow.Select(m => m.Id).ToList();

        SuspendLayout();
        _meetingList.SuspendLayout();
        _meetingList.Controls.Clear();

        foreach (var meeting in meetingsToShow)
        {
            _meetingList.Controls.Add(CreateMeetingRow(meeting, now));
        }

        _meetingList.ResumeLayout();

        Visible = meetingsToShow.Count > 0;
        if (Visible)
        {
            ApplyPreferredSizeAndPosition(meetingsToShow.Count);
            EnsureTopMostWindow();
        }

        UpdateTrayIcon(now);

        ResumeLayout();
    }

    private List<ReminderMeeting> GetMeetingsToShow()
    {
        return _currentMeetings.Take(3).ToList();
    }

    private Control CreateMeetingRow(ReminderMeeting meeting, DateTime now)
    {
        const int RowH = 48;
        var isDismissed = _reminderService.IsDismissed(meeting.Id, now);
        var row = new Panel
        {
            Width = 538,
            Height = RowH,
            Margin = new Padding(0, 0, 0, 4),
            BackColor = GetRowBackColor(meeting, now, isDismissed),
            Tag = meeting
        };

        var accent = new Panel
        {
            Left = 0, Top = 0, Width = 4, Height = RowH,
            BackColor = GetAccentColor(meeting, now, isDismissed)
        };

        // Right-side buttons (Join + dismiss ×)
        var dismissButton = new Button
        {
            Left = 510, Top = 9, Width = 24, Height = 30,
            Text = "×",
            TextAlign = ContentAlignment.MiddleCenter,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(50, 58, 70),
            ForeColor = Color.WhiteSmoke,
            Font = new Font("Segoe UI", 11, FontStyle.Bold)
        };
        dismissButton.FlatAppearance.BorderSize = 0;
        dismissButton.Click += (_, _) =>
        {
            _reminderService.Dismiss(meeting);
            RefreshMeetings(DateTime.Now);
            RebuildView(DateTime.Now);
        };

        var joinButton = new Button
        {
            Left = 456, Top = 9, Width = 50, Height = 30,
            Text = "Join",
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(30, 105, 70),
            ForeColor = Color.WhiteSmoke,
            Visible = meeting.HasTeamsJoinUrl
        };
        joinButton.FlatAppearance.BorderSize = 0;
        joinButton.Click += (_, _) => _reminderService.OpenJoin(meeting);

        int rightEdge = meeting.HasTeamsJoinUrl ? joinButton.Left : dismissButton.Left;
        int countdownWidth = 124;
        int countdownLeft = rightEdge - countdownWidth - 4;

        var countdownLabel = new Label
        {
            Name = "CountdownLabel",
            AutoSize = false,
            Left = countdownLeft, Top = 5,
            Width = countdownWidth, Height = 18,
            TextAlign = ContentAlignment.TopRight,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            ForeColor = GetCountdownColor(meeting, now),
            Text = BuildCountdownText(meeting, now)
        };

        var endCountdownLabel = new Label
        {
            Name = "EndCountdownLabel",
            AutoSize = false,
            Left = countdownLeft, Top = 25,
            Width = countdownWidth, Height = 16,
            TextAlign = ContentAlignment.TopRight,
            Font = new Font("Segoe UI", 8, FontStyle.Regular),
            ForeColor = Color.Gainsboro,
            Text = BuildTimingText(meeting),
            Visible = true
        };

        int textWidth = countdownLeft - 14 - 4;
        var title = new Label
        {
            AutoSize = false,
            AutoEllipsis = true,
            Left = 12, Top = 5,
            Width = textWidth, Height = 18,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            Text = meeting.DisplaySubject
        };

        // Location left, timing lives in the right countdown column
        var locationLabel = new Label
        {
            AutoSize = false,
            AutoEllipsis = true,
            Left = 12, Top = 25,
            Width = textWidth, Height = 16,
            Font = new Font("Segoe UI", 8, FontStyle.Regular),
            ForeColor = Color.Gainsboro,
            Text = BuildLocationText(meeting)
        };

        row.Controls.Add(accent);
        row.Controls.Add(title);
        row.Controls.Add(locationLabel);
        row.Controls.Add(countdownLabel);
        row.Controls.Add(endCountdownLabel);
        row.Controls.Add(joinButton);
        row.Controls.Add(dismissButton);

        return row;
    }

    private void UpdateTrayIcon(DateTime now)
    {
        var next = _nextMeeting;
        if (next is not null && next.End <= now)
        {
            next = null;
        }

        string tooltip;
        if (next is null)
        {
            tooltip = "Meeting Reminder - no upcoming meetings";
        }
        else if (next.IsOngoing(now))
        {
            tooltip = $"Meeting Reminder - ongoing: {next.Subject}";
        }
        else
        {
            var remaining = next.Start - now;
            var timeStr = remaining.TotalHours >= 1
                ? $"in {(int)remaining.TotalHours}h {remaining.Minutes:D2}m"
                : $"in {(int)remaining.TotalMinutes}m";
            tooltip = $"Meeting Reminder - next: {next.Subject} {timeStr}";
        }

        // NotifyIcon tooltip is limited to 63 chars
        _trayIcon.Text = tooltip.Length > 63 ? tooltip[..63] : tooltip;
    }

    private static string BuildLocationText(ReminderMeeting meeting)
    {
        var parts = meeting.Location
            .Split(new[] { '\r', '\n', ';', '|' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(x => x.Trim())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Where(x => !string.Equals(x, "Microsoft Teams Meeting", StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        return parts.Count > 0 ? string.Join(" | ", parts) : "No room";
    }

    private static string BuildTimingText(ReminderMeeting meeting)
    {
        var duration = meeting.End - meeting.Start;
        if (duration < TimeSpan.Zero) duration = TimeSpan.Zero;
        var durationText = duration.TotalHours >= 1
            ? $"{(int)duration.TotalHours}h {duration.Minutes:D2}m"
            : $"{Math.Max(1, (int)Math.Round(duration.TotalMinutes))}m";
        return $"{durationText}, ends {meeting.End:HH:mm}";
    }

    private void ApplyPreferredSizeAndPosition(int visibleRowCount)
    {
        const int rowH = 48;
        const int gap = 4;  // gap between cards
        var newHeight = visibleRowCount * rowH + (visibleRowCount - 1) * gap + 16;
        Height = newHeight;

        var workingArea = Screen.PrimaryScreen?.WorkingArea ?? Screen.FromControl(this).WorkingArea;
        // Sit flush on top of the taskbar, bottom-left
        Left = workingArea.Left + 8;
        Top = workingArea.Bottom - Height; // if only one row, sit above the taskbar; if multiple rows, sit flush with the top of the taskbar
    }

    private void EnsureTopMostWindow()
    {
        if (!IsHandleCreated)
        {
            return;
        }

        SetWindowPos(Handle, HwndTopMost, 0, 0, 0, 0, SwpNoMove | SwpNoSize | SwpNoActivate | SwpShowWindow);
    }

    private void OnDisplaySettingsChanged(object? sender, EventArgs e)
    {
        // Reposition the window when display configuration changes (monitor plug/unplug, etc.)
        if (Visible && _renderedMeetingIds.Count > 0)
        {
            BeginInvoke(() =>
            {
                ApplyPreferredSizeAndPosition(_renderedMeetingIds.Count);
                EnsureTopMostWindow();
            });
        }
    }

    private static Icon CreateTrayBellIcon()
    {
        using var bitmap = new Bitmap(16, 16);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        graphics.Clear(Color.Transparent);

        using var bodyBrush = new SolidBrush(Color.FromArgb(242, 242, 242));
        using var strokePen = new Pen(Color.FromArgb(70, 70, 70), 1f)
        {
            LineJoin = System.Drawing.Drawing2D.LineJoin.Round
        };
        using var clapperBrush = new SolidBrush(Color.FromArgb(70, 70, 70));

        // Minimal bell silhouette: simple dome + skirt + clapper.
        graphics.FillPie(bodyBrush, 3, 3, 10, 9, 180, 180);
        graphics.FillRectangle(bodyBrush, 4, 8, 8, 3);
        graphics.DrawArc(strokePen, 3, 3, 10, 9, 180, 180);
        graphics.DrawLine(strokePen, 4, 8, 4, 11);
        graphics.DrawLine(strokePen, 12, 8, 12, 11);
        graphics.DrawLine(strokePen, 4, 11, 12, 11);
        graphics.FillEllipse(clapperBrush, 7, 11, 2, 2);
        graphics.DrawLine(strokePen, 7, 2, 9, 2);

        var handle = bitmap.GetHicon();
        try
        {
            using var temporaryIcon = Icon.FromHandle(handle);
            return (Icon)temporaryIcon.Clone();
        }
        finally
        {
            DestroyIcon(handle);
        }
    }

    private static Color GetCountdownColor(ReminderMeeting meeting, DateTime now)
    {
        if (meeting.IsOngoing(now) && meeting.IsNotResponded)
        {
            return Color.FromArgb(255, 179, 71);
        }

        if (meeting.IsOngoing(now))
        {
            return Color.FromArgb(255, 216, 107);
        }

        var remaining = meeting.Start - now;
        return remaining <= TimeSpan.FromMinutes(5) ? Color.FromArgb(255, 85, 85) : Color.WhiteSmoke;
    }

    private static string BuildCountdownText(ReminderMeeting meeting, DateTime now)
    {
        if (meeting.IsOngoing(now))
        {
            return "ONGOING";
        }

        var remaining = meeting.Start - now;
        if (remaining < TimeSpan.Zero)
        {
            remaining = TimeSpan.Zero;
        }

        var totalMinutes = (int)remaining.TotalMinutes;
        return $"{totalMinutes:D2}:{remaining.Seconds:D2}";
    }

    private void UpdateRowCountdowns(DateTime now)
    {
        foreach (Control control in _meetingList.Controls)
        {
            if (control is not Panel row || row.Tag is not ReminderMeeting meeting)
            {
                continue;
            }

            if (row.Controls["CountdownLabel"] is Label countdown)
            {
                countdown.Text = BuildCountdownText(meeting, now);
                countdown.ForeColor = GetCountdownColor(meeting, now);
            }

            if (row.Controls["EndCountdownLabel"] is Label endCountdown)
            {
                endCountdown.Visible = true;
                endCountdown.Text = BuildTimingText(meeting);
            }
        }
    }

    private static Color GetRowBackColor(ReminderMeeting meeting, DateTime now, bool isDismissed)
    {
        if (meeting.IsOngoing(now) && meeting.IsNotResponded)
        {
            return Color.FromArgb(66, 52, 37);
        }

        if (meeting.IsOngoing(now))
        {
            return Color.FromArgb(31, 63, 54);
        }

        if (isDismissed)
        {
            return Color.FromArgb(43, 46, 54);
        }

        if (meeting.IsDeclined)
        {
            return Color.FromArgb(57, 40, 43);
        }

        return Color.FromArgb(38, 46, 57);
    }

    private static Color GetAccentColor(ReminderMeeting meeting, DateTime now, bool isDismissed)
    {
        if (meeting.IsOngoing(now) && meeting.IsNotResponded)
        {
            return Color.FromArgb(255, 179, 71);
        }

        if (meeting.IsOngoing(now))
        {
            return Color.FromArgb(72, 201, 142);
        }

        if (isDismissed)
        {
            return Color.FromArgb(150, 156, 166);
        }

        if (meeting.IsDeclined)
        {
            return Color.FromArgb(232, 93, 117);
        }

        if (meeting.IsOverlapping)
        {
            return Color.FromArgb(155, 118, 255);
        }

        return Color.FromArgb(96, 153, 255);
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        Microsoft.Win32.SystemEvents.DisplaySettingsChanged -= OnDisplaySettingsChanged;
        _timer.Stop();
        _timer.Dispose();
        _trayIcon.Visible = false;
        _trayIcon.Dispose();
        _trayBellIcon.Dispose();
        base.OnFormClosed(e);
    }

    [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

    [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
    private static extern bool DestroyIcon(IntPtr hIcon);
}
