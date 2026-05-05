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
    private bool _needsRefresh = true;

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
                if (y > 0) // only if there were past meetings above
                {
                    var sep = CreateNowSeparator(rowWidth);
                    sep.Top = y;
                    _listPanel.Controls.Add(sep);
                    y += sep.Height + 4;
                }
            }

            var row = CreateAgendaRow(meeting, now, rowWidth, isPast);
            row.Top = y;
            _listPanel.Controls.Add(row);
            y += row.Height + 4;
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

        // Re-activate after the COM call to reclaim focus if Outlook briefly stole it.
        if (WindowState == FormWindowState.Normal)
            Activate();
    }

    private static Panel CreateNowSeparator(int rowWidth)
    {
        const int sepH = 20;
        var panel = new Panel
        {
            Left = 0,
            Width = rowWidth,
            Height = sepH,
            BackColor = Color.FromArgb(22, 26, 36)
        };

        panel.Paint += (_, e) =>
        {
            var g = e.Graphics;
            int midY = sepH / 2;
            const string text = "NOW";
            using var font = new Font("Segoe UI", 7, FontStyle.Bold);
            using var brush = new SolidBrush(Color.FromArgb(255, 85, 85));
            using var pen = new Pen(Color.FromArgb(255, 85, 85), 2);
            var textSize = g.MeasureString(text, font);
            float textX = 8;
            float textY = midY - textSize.Height / 2;
            g.DrawString(text, font, brush, textX, textY);
            int lineX = (int)(textX + textSize.Width + 2);
            g.DrawLine(pen, lineX, midY, panel.Width - 4, midY);
        };

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

        int rowH = 54;

        // Build icon slots right-to-left: Chat(💬), Join(📹), Accept(👍), Decline(👎)
        const int iconSize = 28;
        const int iconGap = 2;
        const int rightMargin = 6;
        int cursor = rowWidth - rightMargin;
        int joinLeft = -1, chatLeft = -1, acceptLeft = -1, declineLeft = -1;
        if (hasChat)   { chatLeft    = cursor - iconSize; cursor = chatLeft    - iconGap; }
        if (hasJoin)   { joinLeft    = cursor - iconSize; cursor = joinLeft    - iconGap; }
        if (hasRespond){ declineLeft = cursor - iconSize; cursor = declineLeft - iconGap;
                         acceptLeft  = cursor - iconSize; cursor = acceptLeft  - iconGap; }
        int contentRight = cursor; // right edge for text content

        var rowBg = GetRowBackColor(meeting, now, isPast);
        var row = new Panel
        {
            Left = 0,
            Width = rowWidth,
            Height = rowH,
            BackColor = rowBg
        };

        var accent = new Panel
        {
            Left = 0, Top = 0,
            Width = 4, Height = rowH,
            BackColor = GetAccentColor(meeting, now, isPast)
        };

        var subjectColor = (meeting.IsCancelled || isPast)
            ? Color.FromArgb(120, 120, 130)
            : Color.WhiteSmoke;

        var timeLabel = new Label
        {
            AutoSize = false,
            Left = 12, Top = 9,
            Width = 88, Height = 18,
            Font = new Font("Segoe UI", 8, FontStyle.Regular),
            ForeColor = Color.FromArgb(180, 180, 195),
            Text = $"{meeting.Start:HH:mm}–{meeting.End:HH:mm}",
            TextAlign = ContentAlignment.MiddleLeft
        };

        // Account label sits just left of the icon cluster
        const int accountWidth = 110;
        bool showAccount = !string.IsNullOrEmpty(meeting.Account);
        int accountLeft = contentRight - accountWidth;
        var accountLabel = new Label
        {
            AutoSize = false,
            AutoEllipsis = true,
            Left = Math.Max(104, accountLeft), Top = 10,
            Width = accountWidth, Height = 16,
            Font = new Font("Segoe UI", 7, FontStyle.Regular),
            ForeColor = Color.FromArgb(120, 125, 140),
            Text = meeting.Account.Contains('@') ? meeting.Account[(meeting.Account.IndexOf('@') + 1)..] : meeting.Account,
            TextAlign = ContentAlignment.MiddleRight,
            Visible = showAccount
        };

        int subjectRight = showAccount ? accountLeft - 4 : contentRight - 4;
        var subjectLabel = new Label
        {
            AutoSize = false,
            AutoEllipsis = true,
            Left = 104, Top = 8,
            Width = Math.Max(0, subjectRight - 104),
            Height = 20,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            ForeColor = subjectColor,
            Text = meeting.DisplaySubject,
            TextAlign = ContentAlignment.MiddleLeft
        };

        var locationLabel = new Label
        {
            AutoSize = false,
            AutoEllipsis = true,
            Left = 12, Top = 28,
            Width = Math.Max(0, contentRight - 24),
            Height = 18,
            Font = new Font("Segoe UI", 8, FontStyle.Regular),
            ForeColor = Color.FromArgb(160, 165, 180),
            Text = BuildLocationText(meeting),
            TextAlign = ContentAlignment.MiddleLeft
        };

        int iconTop = (54 - iconSize) / 2;

        row.Controls.Add(accent);
        row.Controls.Add(timeLabel);
        row.Controls.Add(subjectLabel);
        row.Controls.Add(accountLabel);
        row.Controls.Add(locationLabel);

        if (hasJoin)
        {
            var joinBtn = MaterialIcons.MakeButton(MaterialIcons.VideoCall, joinLeft, iconTop, iconSize,
                Color.FromArgb(0, 200, 83), rowBg, 0.78f);
            joinBtn.Click += (_, _) => _reminderService.OpenJoin(meeting);
            row.Controls.Add(joinBtn);
        }

        if (hasChat)
        {
            var chatBtn = MaterialIcons.MakeButton(MaterialIcons.ChatBubble, chatLeft, iconTop, iconSize, Color.FromArgb(100, 160, 230), rowBg);
            chatBtn.Click += (_, _) => Process.Start(new ProcessStartInfo
            {
                FileName = meeting.TeamsChatUrl!,
                UseShellExecute = true
            });
            row.Controls.Add(chatBtn);
        }

        if (hasRespond)
        {
            var acceptBtn = MaterialIcons.MakeButton(MaterialIcons.ThumbUp, acceptLeft, iconTop, iconSize, Color.FromArgb(80, 190, 120), rowBg);
            acceptBtn.Click += (_, _) =>
            {
                try { _reminderService.RespondToMeeting(meeting.Id, true); }
                catch { }
                RefreshAgenda();
            };
            row.Controls.Add(acceptBtn);

            var declineBtn = MaterialIcons.MakeButton(MaterialIcons.ThumbDown, declineLeft, iconTop, iconSize, Color.FromArgb(210, 80, 90), rowBg);
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

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        base.OnFormClosed(e);
        _ownerHandle.Dispose();
    }
}
