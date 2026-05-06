using System.Diagnostics;
using System.Runtime.Versioning;

namespace Outlook.ReminderApp;

[SupportedOSPlatform("windows")]
internal sealed class AgendaForm : Form
{
    private const int WsExAppWindow = 0x00040000;
    private const int WsExToolWindow = 0x00000080;
    private const int WmSysCommand = 0x0112;
    private const int ScRestore = 0xF120;

    // A hidden WS_EX_TOOLWINDOW native window used as the owner of AgendaForm.
    // Owned windows are excluded from Alt+Tab; WS_EX_APPWINDOW on AgendaForm
    // still forces a taskbar button. The owner itself is invisible and off-screen.
    private sealed class ToolWindowHandle : NativeWindow, IDisposable
    {
        public ToolWindowHandle()
        {
            CreateHandle(new CreateParams
            {
                Caption = string.Empty,
                Style = unchecked((int)0x80000000), // WS_POPUP
                ExStyle = WsExToolWindow,
                Width = 0,
                Height = 0
            });
        }

        public void Dispose() => DestroyHandle();
    }

    private readonly ToolWindowHandle _ownerHandle = new();
    private readonly MeetingReminderService _reminderService;
    private readonly MeetingCache _cache;
    private Panel _listPanel = null!;
    private System.Windows.Forms.Timer _countdownTimer = null!;
    private bool _needsRefresh = true;
    private Panel? _nowSepPanel;
    private Label? _nowSepCountdownLabel;
    private Color _nowSepColor = Color.FromArgb(255, 85, 85);

    public AgendaForm(MeetingReminderService reminderService, MeetingCache cache)
    {
        _reminderService = reminderService;
        _cache = cache;

        AutoScaleMode = AutoScaleMode.None;
        FormBorderStyle = FormBorderStyle.None;
        StartPosition = FormStartPosition.Manual;
        ShowInTaskbar = true;
        Text = "Today's Agenda";
        TopMost = false;
        BackColor = Color.FromArgb(22, 26, 36);
        ForeColor = Color.WhiteSmoke;
        Width = 480;
        Height = 400;

        // Pre-center with initial size so the first restore appears in the right place.
        var screenArea = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1920, 1080);
        Left = screenArea.Left + (screenArea.Width - Width) / 2;
        Top = screenArea.Top + (screenArea.Height - Height) / 2;

        // Start minimized so Show() puts us in the taskbar without going through Normal.
        WindowState = FormWindowState.Minimized;

        var appIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        if (appIcon is not null) Icon = appIcon;

        BuildLayout();

        _countdownTimer = new System.Windows.Forms.Timer { Interval = 1000 };
        _countdownTimer.Tick += (_, _) => UpdateCountdown();
        _countdownTimer.Start();

        // SizeChanged fires reliably when the window is restored from Minimized → Normal.
        SizeChanged += (_, _) =>
        {
            if (WindowState == FormWindowState.Normal && _needsRefresh)
            {
                _needsRefresh = false;
                RefreshAgenda();
            }
        };

        // Fallback: also handle Activated in case the window was already Normal
        // and regains focus without a size change (e.g. Alt+Tab back).
        Activated += (_, _) =>
        {
            if (WindowState == FormWindowState.Normal && _needsRefresh)
            {
                _needsRefresh = false;
                RefreshAgenda();
            }
        };

        Deactivate += (_, _) =>
        {
            BeginInvoke(() =>
            {
                WindowState = FormWindowState.Minimized;
                _needsRefresh = true;
            });
        };

        _cache.Refreshed += (_, _) =>
        {
            if (WindowState == FormWindowState.Normal)
                RefreshAgenda();
        };
    }

    // Set Opacity=0 before the restore paints so the window appears fully built
    // in a single frame (layout rebuild happens while invisible).
    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WmSysCommand && ((int)m.WParam & 0xFFF0) == ScRestore && IsHandleCreated)
        {
            Opacity = 0;
            base.WndProc(ref m); // triggers WM_SIZE → SizeChanged → RefreshAgenda
            Opacity = 1;
            return;
        }
        base.WndProc(ref m);
    }

    protected override CreateParams CreateParams
    {
        get
        {
            var cp = base.CreateParams;
            cp.ExStyle |= WsExAppWindow;
            cp.ExStyle &= ~WsExToolWindow;
            // Setting Parent on a WS_POPUP window sets the owner (not the visual parent).
            // This combined with WS_EX_APPWINDOW gives: taskbar button yes, Alt+Tab no.
            cp.Parent = _ownerHandle.Handle;
            return cp;
        }
    }

    private void BuildLayout()
    {
        var header = new Panel
        {
            Dock = DockStyle.Top,
            Height = 36,
            BackColor = Color.FromArgb(30, 34, 44)
        };

        var dateLabel = new Label
        {
            AutoSize = false,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(12, 0, 12, 0),
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = Color.WhiteSmoke,
            Text = DateTime.Now.ToString("dddd, d MMMM")
        };

        header.Controls.Add(dateLabel);

        _listPanel = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            Padding = Padding.Empty,
            BackColor = Color.FromArgb(22, 26, 36)
        };

        Controls.Add(_listPanel);
        Controls.Add(header);
    }

    private void CenterOnScreen()
    {
        var screen = Screen.PrimaryScreen?.WorkingArea ?? Screen.FromControl(this).WorkingArea;
        Left = screen.Left + (screen.Width - Width) / 2;
        Top = screen.Top + (screen.Height - Height) / 2;
    }

    private void RefreshAgenda()
    {
        var now = DateTime.Now;
        var meetings = _reminderService.GetTodaysMeetings(now, _cache.All);

        SuspendLayout();
        _listPanel.SuspendLayout();

        _nowSepPanel = null;
        _nowSepCountdownLabel = null;

        foreach (Control c in _listPanel.Controls)
            c.Dispose();
        _listPanel.Controls.Clear();

        const int rowWidth = 480;

        int y = 0;
        bool separatorInserted = false;
        foreach (var meeting in meetings)
        {
            bool isPast = meeting.End <= now && !meeting.IsOngoing(now);

            // Insert "now" separator before the first ongoing or future meeting
            if (!separatorInserted && !isPast)
            {
                separatorInserted = true;
                var sep = CreateNowSeparator(rowWidth);
                sep.Top = y;
                _listPanel.Controls.Add(sep);
                y += sep.Height + 4;
            }

            var row = CreateAgendaRow(meeting, now, rowWidth, isPast);
            row.Top = y;
            _listPanel.Controls.Add(row);
            y += row.Height + 4;
        }

        // If all meetings were in the past (or no meetings), append the NOW separator at bottom
        if (!separatorInserted)
        {
            var sep = CreateNowSeparator(rowWidth);
            sep.Top = y;
            _listPanel.Controls.Add(sep);
            y += sep.Height + 4;
        }

        if (meetings.Count == 0)
        {
            var emptyLabel = new Label
            {
                AutoSize = false,
                Left = 12, Top = 16,
                Width = rowWidth, Height = 32,
                Text = "No meetings today",
                Font = new Font("Segoe UI", 10, FontStyle.Regular),
                ForeColor = Color.FromArgb(140, 145, 160),
                TextAlign = ContentAlignment.MiddleCenter
            };
            _listPanel.Controls.Add(emptyLabel);
            y = 60;
        }

        int maxHeight = (Screen.PrimaryScreen?.WorkingArea.Height ?? 768) * 3 / 4;
        int contentHeight = y + 36; // rows + header
        Height = Math.Min(Math.Max(contentHeight, 100), maxHeight);

        CenterOnScreen();

        _listPanel.ResumeLayout();
        ResumeLayout();

        UpdateCountdown();

        // Re-activate after the COM call to reclaim focus if Outlook briefly stole it.
        if (WindowState == FormWindowState.Normal)
            Activate();
    }

    private Panel CreateNowSeparator(int rowWidth)
    {
        const int sepH = 22;
        const int countdownWidth = 150;

        var panel = new Panel
        {
            Left = 0,
            Width = rowWidth,
            Height = sepH,
            BackColor = Color.FromArgb(22, 26, 36)
        };

        var countdownLbl = new Label
        {
            AutoSize = false,
            Left = rowWidth - countdownWidth - 6,
            Top = 0,
            Width = countdownWidth,
            Height = sepH,
            Font = new Font("Segoe UI", 7.5f, FontStyle.Bold),
            ForeColor = _nowSepColor,
            TextAlign = ContentAlignment.MiddleRight,
            Text = string.Empty
        };

        panel.Paint += (_, e) =>
        {
            var g = e.Graphics;
            int midY = sepH / 2;
            const string text = "NOW";
            using var font = new Font("Segoe UI", 7, FontStyle.Bold);
            using var brush = new SolidBrush(_nowSepColor);
            using var pen = new Pen(_nowSepColor, 2);
            var textSize = g.MeasureString(text, font);
            float textX = 8;
            float textY = midY - textSize.Height / 2;
            g.DrawString(text, font, brush, textX, textY);
            int lineX = (int)(textX + textSize.Width + 2);
            int lineEnd = countdownLbl.Left - 4;
            g.DrawLine(pen, lineX, midY, lineEnd, midY);
        };

        panel.Controls.Add(countdownLbl);
        _nowSepPanel = panel;
        _nowSepCountdownLabel = countdownLbl;
        return panel;
    }

    private Panel CreateAgendaRow(ReminderMeeting meeting, DateTime now, int rowWidth, bool isPast = false)
    {
        bool hasJoin = meeting.HasTeamsJoinUrl;
        bool hasChat = meeting.TeamsChatUrl is not null;
        bool hasRespond = meeting.IsResponseRequested
            && !meeting.IsCancelled
            && !string.Equals(meeting.ResponseStatus, "Accepted", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(meeting.ResponseStatus, "Declined", StringComparison.OrdinalIgnoreCase);

        bool hasIcons = hasJoin || hasChat || hasRespond;
        bool showAccount = !string.IsNullOrEmpty(meeting.Account);

        const int iconSize     = 28;
        const int iconGap      = 2;
        const int rightMargin  = 6;
        const int leftColLeft  = 12;
        const int leftColWidth = 88;
        const int midColLeft   = 104;

        // Row 1 right column: account label (fixed width)
        const int accountWidth = 120;
        int row1RightLeft  = rowWidth - rightMargin - accountWidth;
        int midColWidthRow1 = Math.Max(0, row1RightLeft - midColLeft - 4);

        // Row 2 right column: icon cluster (sized to actual icon count)
        int iconCount = (hasJoin ? 1 : 0) + (hasChat ? 1 : 0) + (hasRespond ? 2 : 0);
        int iconClusterWidth = iconCount > 0 ? iconCount * iconSize + (iconCount - 1) * iconGap : 0;
        int row2RightLeft = rowWidth - rightMargin - iconClusterWidth;
        int midColWidthRow2 = hasIcons ? Math.Max(0, row2RightLeft - midColLeft - 4) : Math.Max(0, rowWidth - rightMargin - midColLeft - 4);

        const int line1Top = 6;
        const int line1H   = 18;
        const int line2Top = 28;
        int line2H = hasIcons ? iconSize : 17;
        int rowH   = line2Top + line2H + 4;

        // Duration string
        var duration = meeting.End - meeting.Start;
        string durationText = duration.TotalHours >= 1
            ? (duration.Minutes > 0 ? $"{(int)duration.TotalHours}h {duration.Minutes}m" : $"{(int)duration.TotalHours}h")
            : $"{(int)duration.TotalMinutes} min";

        var rowBg = GetRowBackColor(meeting, now, isPast);
        var row = new Panel { Left = 0, Width = rowWidth, Height = rowH, BackColor = rowBg };
        var accent = new Panel { Left = 0, Top = 0, Width = 4, Height = rowH, BackColor = GetAccentColor(meeting, now, isPast) };

        var subjectColor = (meeting.IsCancelled || isPast)
            ? Color.FromArgb(120, 120, 130) : Color.WhiteSmoke;

        // ── Row 1: time | subject | account ──
        var timeLabel = new Label
        {
            AutoSize = false,
            Left = leftColLeft, Top = line1Top, Width = leftColWidth, Height = line1H,
            Font = new Font("Segoe UI", 8, FontStyle.Regular),
            ForeColor = Color.FromArgb(180, 180, 195),
            Text = $"{meeting.Start:HH:mm}–{meeting.End:HH:mm}",
            TextAlign = ContentAlignment.MiddleLeft
        };

        var subjectLabel = new Label
        {
            AutoSize = false, AutoEllipsis = true,
            Left = midColLeft, Top = line1Top, Width = midColWidthRow1, Height = line1H,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            ForeColor = subjectColor,
            Text = meeting.DisplaySubject,
            TextAlign = ContentAlignment.MiddleLeft
        };

        var accountLabel = new Label
        {
            AutoSize = false, AutoEllipsis = true,
            Left = row1RightLeft, Top = line1Top, Width = accountWidth, Height = line1H,
            Font = new Font("Segoe UI", 7, FontStyle.Regular),
            ForeColor = Color.FromArgb(110, 115, 135),
            Text = meeting.Account.Contains('@') ? meeting.Account[(meeting.Account.IndexOf('@') + 1)..] : meeting.Account,
            TextAlign = ContentAlignment.MiddleRight,
            Visible = showAccount
        };

        // ── Row 2: duration | location | icons ──
        var durationLabel = new Label
        {
            AutoSize = false,
            Left = leftColLeft, Top = line2Top, Width = leftColWidth, Height = line2H,
            Font = new Font("Segoe UI", 7.5f, FontStyle.Regular),
            ForeColor = Color.FromArgb(110, 115, 135),
            Text = durationText,
            TextAlign = ContentAlignment.MiddleLeft
        };

        var locationLabel = new Label
        {
            AutoSize = false, AutoEllipsis = true,
            Left = midColLeft, Top = line2Top, Width = midColWidthRow2, Height = line2H,
            Font = new Font("Segoe UI", 8, FontStyle.Regular),
            ForeColor = Color.FromArgb(160, 165, 180),
            Text = BuildLocationText(meeting),
            TextAlign = ContentAlignment.MiddleLeft
        };

        row.Controls.Add(accent);
        row.Controls.Add(timeLabel);
        row.Controls.Add(subjectLabel);
        row.Controls.Add(accountLabel);
        row.Controls.Add(durationLabel);
        row.Controls.Add(locationLabel);

        if (!hasIcons) return row;

        // Icons right-to-left, starting from row2RightLeft
        int iconTop = line2Top;
        int cursor = rowWidth - rightMargin;
        int joinLeft = -1, chatLeft = -1, acceptLeft = -1, declineLeft = -1;
        if (hasChat)    { chatLeft    = cursor - iconSize; cursor = chatLeft    - iconGap; }
        if (hasJoin)    { joinLeft    = cursor - iconSize; cursor = joinLeft    - iconGap; }
        if (hasRespond) { declineLeft = cursor - iconSize; cursor = declineLeft - iconGap;
                          acceptLeft  = cursor - iconSize; }

        if (hasJoin)
        {
            var joinBtn = MaterialIcons.MakeButton(MaterialIcons.VideoCall, joinLeft, iconTop, iconSize,
                Color.FromArgb(0, 200, 83), rowBg, 0.78f);
            joinBtn.Click += (_, _) => _reminderService.OpenJoin(meeting);
            row.Controls.Add(joinBtn);
        }

        if (hasChat)
        {
            var chatBtn = MaterialIcons.MakeButton(MaterialIcons.ChatBubble, chatLeft, iconTop, iconSize,
                Color.FromArgb(100, 160, 230), rowBg);
            chatBtn.Click += (_, _) => Process.Start(new ProcessStartInfo
            {
                FileName = meeting.TeamsChatUrl!,
                UseShellExecute = true
            });
            row.Controls.Add(chatBtn);
        }

        if (hasRespond)
        {
            var acceptBtn = MaterialIcons.MakeButton(MaterialIcons.ThumbUp, acceptLeft, iconTop, iconSize,
                Color.FromArgb(80, 190, 120), rowBg);
            acceptBtn.Click += (_, _) =>
            {
                try { _reminderService.RespondToMeeting(meeting.Id, true); }
                catch { }
                RefreshAgenda();
            };
            row.Controls.Add(acceptBtn);

            var declineBtn = MaterialIcons.MakeButton(MaterialIcons.ThumbDown, declineLeft, iconTop, iconSize,
                Color.FromArgb(210, 80, 90), rowBg);
            declineBtn.Click += (_, _) =>
            {
                try { _reminderService.RespondToMeeting(meeting.Id, false); }
                catch { }
                RefreshAgenda();
            };
            row.Controls.Add(declineBtn);
        }

        return row;
    }

    private static Color GetRowBackColor(ReminderMeeting meeting, DateTime now, bool isPast = false)
    {
        if (isPast) return Color.FromArgb(26, 29, 38);
        if (meeting.IsCancelled)
            return Color.FromArgb(57, 40, 43);
        if (meeting.IsOngoing(now) && meeting.IsNotResponded)
            return Color.FromArgb(66, 52, 37);
        if (meeting.IsOngoing(now))
            return Color.FromArgb(31, 63, 54);
        if (meeting.IsDeclined)
            return Color.FromArgb(57, 40, 43);
        return Color.FromArgb(38, 46, 57);
    }

    private static Color GetAccentColor(ReminderMeeting meeting, DateTime now, bool isPast = false)
    {
        if (isPast) return Color.FromArgb(55, 58, 70);
        if (meeting.IsCancelled)
            return Color.FromArgb(232, 93, 117);
        if (meeting.IsOngoing(now) && meeting.IsNotResponded)
            return Color.FromArgb(255, 179, 71);
        if (meeting.IsOngoing(now))
            return Color.FromArgb(72, 201, 142);
        if (meeting.IsDeclined)
            return Color.FromArgb(232, 93, 117);
        if (meeting.IsNotResponded)
            return Color.FromArgb(255, 179, 71);
        return Color.FromArgb(96, 153, 255);
    }

    private static string BuildLocationText(ReminderMeeting meeting)
    {
        if (meeting.IsCancelled)
            return "Cancelled";

        var parts = meeting.Location
            .Split(['\r', '\n', ';', '|'], StringSplitOptions.RemoveEmptyEntries)
            .Select(x => x.Trim())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Where(x => !string.Equals(x, "Microsoft Teams Meeting", StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        return parts.Count > 0 ? string.Join(" | ", parts) : (meeting.IsMeeting ? "Online" : "No location");
    }

    private void UpdateCountdown()
    {
        var now = DateTime.Now;
        var next = _cache.All
            .Where(m => !m.IsCancelled && m.Start.Date == now.Date && m.Start > now)
            .OrderBy(m => m.Start)
            .FirstOrDefault();

        string text;
        Color color;

        if (next is null)
        {
            text = "no more meetings";
            color = Color.FromArgb(80, 85, 100);
        }
        else
        {
            var diff = next.Start - now;
            if (diff.TotalMinutes >= 60)
            {
                int hours = (int)diff.TotalHours;
                int minutes = diff.Minutes;
                text = minutes > 0 ? $"in {hours}h {minutes}m" : $"in {hours}h";
                color = Color.FromArgb(80, 190, 120); // green
            }
            else if (diff.TotalMinutes >= 5)
            {
                int minutes = (int)Math.Ceiling(diff.TotalMinutes);
                text = $"in {minutes} min";
                color = Color.FromArgb(255, 195, 60); // yellow
            }
            else
            {
                int totalSeconds = Math.Max(0, (int)Math.Ceiling(diff.TotalSeconds));
                int m = totalSeconds / 60;
                int s = totalSeconds % 60;
                text = $"starting {m}:{s:D2}";
                color = Color.FromArgb(255, 85, 85); // red
            }
        }

        bool colorChanged = _nowSepColor != color;
        _nowSepColor = color;
        if (_nowSepCountdownLabel is not null)
        {
            _nowSepCountdownLabel.Text = text;
            _nowSepCountdownLabel.ForeColor = color;
        }
        if (colorChanged)
            _nowSepPanel?.Invalidate();
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        base.OnFormClosed(e);
        _countdownTimer.Dispose();
        _ownerHandle.Dispose();
    }
}
