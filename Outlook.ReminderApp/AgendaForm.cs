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
    private const int FadeDurationMs = 100;
    private const int FadeTickIntervalMs = 20;
    private const int MaxNavigationDays = 7;
    private static readonly TimeSpan StaleSelectionResetThreshold = TimeSpan.FromMinutes(5);

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
    private readonly SyncUiController _syncUi;
    private Panel _listPanel = null!;
    private System.Windows.Forms.Timer _countdownTimer = null!;
    private Label? _staleLabel;
    private Label _dateLabel = null!;
    private Label _dayOffsetLabel = null!;
    private DateTime _selectedDate = DateTime.Today;
    private bool _needsRefresh = true;
    private Panel? _nowSepPanel;
    private Label? _nowSepCountdownLabel;
    private Color _nowSepColor = Color.FromArgb(255, 85, 85);
    private MeetingDetailForm? _activeDetailForm;
    private bool _needsInitialCenter = true;
    private string _lastMeetingsFingerprint = string.Empty;
    private string _lastLoadingMessage = string.Empty;
    private int _screenCheckTick;
    private Screen? _preferredScreen;
    private System.Windows.Forms.Timer? _fadeTimer;
    private DateTime _fadeStartUtc;
    private Label _prevButton = null!;
    private Label _nextButton = null!;
    private string _nowSepLabelText = "NOW";
    private DateTime? _minimizedAtUtc;

    public AgendaForm(MeetingReminderService reminderService, MeetingCache cache, SyncUiController syncUi)
    {
        _reminderService = reminderService;
        _cache = cache;
        _syncUi = syncUi;

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
        var screenArea = Screen.FromPoint(Cursor.Position).WorkingArea;
        Left = screenArea.Left + (screenArea.Width - Width) / 2;
        Top = screenArea.Top + (screenArea.Height - Height) / 2;

        // Start minimized so Show() puts us in the taskbar without going through Normal.
        WindowState = FormWindowState.Minimized;

        var appIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        if (appIcon is not null) Icon = appIcon;

        ContextMenuStrip = BuildContextMenu();

        BuildLayout();

        _countdownTimer = new System.Windows.Forms.Timer { Interval = 1000 };
        _countdownTimer.Tick += (_, _) =>
        {
            UpdateCountdown();
            if (++_screenCheckTick >= 30)
            {
                _screenCheckTick = 0;
                PeriodicScreenCheck();
            }
        };
        _countdownTimer.Start();

        Microsoft.Win32.SystemEvents.DisplaySettingsChanged += OnDisplaySettingsChanged;

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
            // Capture cursor screen NOW — before the cursor moves to the taskbar.
            var screenAtDeactivate = Screen.FromPoint(Cursor.Position);
            BeginInvoke(() =>
            {
                // Don't minimize if the detail flyout just grabbed focus
                if (_activeDetailForm is { IsDisposed: false } f &&
                    (f == Form.ActiveForm || f.ContainsFocus))
                    return;

                _activeDetailForm?.Close();
                _activeDetailForm = null;
                _preferredScreen = screenAtDeactivate;
                WindowState = FormWindowState.Minimized;
                _needsRefresh = true;
                _needsInitialCenter = true;
                _minimizedAtUtc = DateTime.UtcNow;
            });
        };

        _cache.Refreshed += (_, _) =>
        {
            if (WindowState == FormWindowState.Normal)
                RefreshAgenda();
        };
    }

    private ContextMenuStrip BuildContextMenu()
    {
        var menu = new ContextMenuStrip();
        menu.Items.Add("Sync settings...", null, (_, _) => _syncUi.ShowConfig(this));
        return menu;
    }

    // Set Opacity=0 before the restore paints so the window appears fully built
    // in a single frame (layout rebuild happens while invisible), then fade in.
    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WmSysCommand && ((int)m.WParam & 0xFFF0) == ScRestore && IsHandleCreated)
        {
            // The cursor is right on the taskbar button being clicked at this point,
            // so it's a fresher/better signal for the target screen than whatever
            // _preferredScreen was left over from the last deactivate.
            _preferredScreen = Screen.FromPoint(Cursor.Position);
            _needsInitialCenter = true;

            if (_minimizedAtUtc is { } minimizedAt &&
                DateTime.UtcNow - minimizedAt >= StaleSelectionResetThreshold &&
                _selectedDate.Date != DateTime.Today)
            {
                ApplySelectedDate(DateTime.Today);
            }
            _minimizedAtUtc = null;

            Opacity = 0;
            base.WndProc(ref m); // triggers WM_SIZE → SizeChanged → RefreshAgenda
            StartFadeIn();
            return;
        }
        base.WndProc(ref m);
    }

    private void StartFadeIn()
    {
        _fadeTimer?.Stop();
        _fadeTimer?.Dispose();

        _fadeStartUtc = DateTime.UtcNow;
        _fadeTimer = new System.Windows.Forms.Timer { Interval = FadeTickIntervalMs };
        _fadeTimer.Tick += (_, _) =>
        {
            var elapsedMs = (DateTime.UtcNow - _fadeStartUtc).TotalMilliseconds;
            var progress = Math.Clamp(elapsedMs / FadeDurationMs, 0, 1);
            // Cubic ease-out keeps the same duration while reducing perceived pop-in.
            var eased = 1 - Math.Pow(1 - progress, 3);
            Opacity = eased;

            if (progress >= 1)
            {
                _fadeTimer.Stop();
                _fadeTimer.Dispose();
                _fadeTimer = null;
            }
        };
        _fadeTimer.Start();
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

        _dateLabel = new Label
        {
            AutoSize = false,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(12, 0, 0, 0),
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = Color.WhiteSmoke,
            Text = _selectedDate.ToString("dddd, d MMMM")
        };

        var navPanel = new Panel
        {
            Dock = DockStyle.Right,
            Width = 100,
            BackColor = Color.Transparent
        };

        _prevButton = MaterialIcons.MakeButton(MaterialIcons.ChevronLeft, 0, 4, 28,
            Color.FromArgb(140, 145, 160), Color.FromArgb(30, 34, 44));
        _prevButton.Click += (_, _) => ChangeSelectedDate(-1);
        navPanel.Controls.Add(_prevButton);

        _dayOffsetLabel = new Label
        {
            Left = 28, Top = 4, Width = 44, Height = 28,
            Font = new Font("Segoe UI", 8, FontStyle.Bold),
            ForeColor = Color.FromArgb(140, 145, 160),
            TextAlign = ContentAlignment.MiddleCenter,
            Text = string.Empty,
            Visible = false
        };
        navPanel.Controls.Add(_dayOffsetLabel);

        _nextButton = MaterialIcons.MakeButton(MaterialIcons.ChevronRight, 72, 4, 28,
            Color.FromArgb(140, 145, 160), Color.FromArgb(30, 34, 44));
        _nextButton.Click += (_, _) => ChangeSelectedDate(1);
        navPanel.Controls.Add(_nextButton);

        header.Controls.Add(_dateLabel);
        header.Controls.Add(navPanel);

        _staleLabel = new Label
        {
            AutoSize = false,
            Dock = DockStyle.Right,
            Width = 200,
            TextAlign = ContentAlignment.MiddleRight,
            Padding = new Padding(0, 0, 12, 0),
            Font = new Font("Segoe UI", 8, FontStyle.Italic),
            ForeColor = Color.FromArgb(255, 180, 80),
            Text = "⚠ Showing cached data",
            Visible = false
        };
        header.Controls.Add(_staleLabel);

        _listPanel = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = false,
            Padding = Padding.Empty,
            BackColor = Color.FromArgb(22, 26, 36)
        };

        var loadingLabel = new Label
        {
            AutoSize = false,
            Dock = DockStyle.Fill,
            Text = "Loading...",
            Font = new Font("Segoe UI", 10, FontStyle.Regular),
            ForeColor = Color.FromArgb(140, 145, 160),
            TextAlign = ContentAlignment.MiddleCenter
        };
        _listPanel.Controls.Add(loadingLabel);

        Controls.Add(_listPanel);
        Controls.Add(header);
    }

    private static string ComputeMeetingsFingerprint(IReadOnlyList<ReminderMeeting> meetings, DateTime now)
    {
        if (meetings.Count == 0) return "empty";
        var sb = new System.Text.StringBuilder(meetings.Count * 64);
        foreach (var m in meetings)
        {
            bool isPast = m.End <= now && !m.IsOngoing(now);
            sb.Append(m.Id).Append('|')
              .Append(m.Subject).Append('|')
              .Append(m.Start.Ticks).Append('|')
              .Append(m.End.Ticks).Append('|')
              .Append(m.ResponseStatus).Append('|')
              .Append(m.IsCancelled).Append('|')
              .Append(m.IsMeeting).Append('|')
              .Append(m.IsAllDay).Append('|')
              .Append(m.IsOverlapping).Append('|')
              .Append(m.TeamsJoinUrl).Append('|')
              .Append(m.Location).Append('|')
              .Append(m.Account).Append('|')
              .Append(isPast).Append(';');
        }
        return sb.ToString();
    }

    private Rectangle GetTargetWorkingArea()
    {
        // Use the screen the cursor was on when the form was last deactivated
        // (at that point the cursor is on the working screen, not the taskbar).
        // Fall back to cursor-current screen for the very first open.
        return (_preferredScreen ?? Screen.FromPoint(Cursor.Position)).WorkingArea;
    }

    private void CenterOnScreen()
    {
        var workingArea = GetTargetWorkingArea();
        Left = workingArea.Left + (workingArea.Width - Width) / 2;
        Top = workingArea.Top + (workingArea.Height - Height) / 2;
    }

    private void OnDisplaySettingsChanged(object? sender, EventArgs e)
    {
        if (WindowState == FormWindowState.Normal)
            BeginInvoke(CenterOnScreen);
    }

    private void PeriodicScreenCheck()
    {
        if (WindowState != FormWindowState.Normal) return;
        // If the screen the form is currently on no longer exists, reposition to cursor screen.
        var currentScreen = Screen.FromControl(this);
        bool stillValid = Screen.AllScreens.Any(s => s.Bounds.Equals(currentScreen.Bounds));
        if (!stillValid)
            BeginInvoke(CenterOnScreen);
    }

    private void RefreshAgenda()
    {
        var now = DateTime.Now;
        bool shouldCenterOnShow = _needsInitialCenter;
        _needsInitialCenter = false;

        // Toggle stale-data indicator in the header
        if (_staleLabel is not null)
        {
            if (!_cache.IsLoaded && _cache.LastRefreshFailed)
            {
                _staleLabel.Text = "⚠ Outlook not responding";
                _staleLabel.Visible = true;
            }
            else if (_cache.LastRefreshFailed)
            {
                _staleLabel.Text = "⚠ Showing cached data";
                _staleLabel.Visible = true;
            }
            else
            {
                _staleLabel.Visible = false;
            }
        }

        if (!_cache.IsLoaded)
        {
            if (shouldCenterOnShow)
                CenterOnScreen();

            // Data not yet available — ensure loading/error label is visible and wait for Refreshed event.
            var message = _cache.LastRefreshFailed && !string.IsNullOrEmpty(_cache.LastError)
                ? $"Could not load Outlook data:\n{_cache.LastError}"
                : "Loading...";

            if (_listPanel.Controls.Count != 1 ||
                _listPanel.Controls[0] is not Label lbl ||
                lbl.Text != message)
            {
                foreach (Control c in _listPanel.Controls) c.Dispose();
                _listPanel.Controls.Clear();
                _listPanel.Controls.Add(new Label
                {
                    AutoSize = false,
                    Dock = DockStyle.Fill,
                    Text = message,
                    Font = new Font("Segoe UI", 10, FontStyle.Regular),
                    ForeColor = _cache.LastRefreshFailed
                        ? Color.FromArgb(255, 180, 80)
                        : Color.FromArgb(140, 145, 160),
                    TextAlign = ContentAlignment.MiddleCenter
                });
            }
            _lastLoadingMessage = message;
            return;
        }

        _lastLoadingMessage = string.Empty;

        var dayStart = _selectedDate.Date;
        var dayEnd   = dayStart.AddDays(1);

        var meetings = _cache.All
            .Where(m => !m.IsOutlookSynced) // hide busy-block placeholders created by outlook-sync
            .Where(m => m.Start < dayEnd && m.End > dayStart) // overlap with selected day
            .OrderBy(m => m.Start)
            .ThenBy(m => m.IsAllDay ? 0 : 1) // all-day events first on their start day
            .ToList();

        var fingerprint = ComputeMeetingsFingerprint(meetings, now);
        if (fingerprint == _lastMeetingsFingerprint)
        {
            if (shouldCenterOnShow)
                CenterOnScreen();

            UpdateCountdown();
            return;
        }
        _lastMeetingsFingerprint = fingerprint;

        SuspendLayout();
        _listPanel.SuspendLayout();

        _nowSepPanel = null;
        _nowSepCountdownLabel = null;

        // Close any open detail flyout before rebuilding rows
        _activeDetailForm?.Close();
        _activeDetailForm = null;

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
                Text = _selectedDate.Date == DateTime.Today ? "No meetings today" : "No meetings on this day",
                Font = new Font("Segoe UI", 10, FontStyle.Regular),
                ForeColor = Color.FromArgb(140, 145, 160),
                TextAlign = ContentAlignment.MiddleCenter
            };
            _listPanel.Controls.Add(emptyLabel);
            y = 60;
        }

        int maxHeight = GetTargetWorkingArea().Height * 3 / 4;
        int contentHeight = y + 36; // rows + header
        Height = Math.Min(Math.Max(contentHeight, 100), maxHeight);

        // Enable AutoScroll only if content exceeds visible area
        _listPanel.AutoScroll = (y > Height - 36);

        if (shouldCenterOnShow)
            CenterOnScreen();

        // Keep nav buttons in sync with the current position (handles midnight rollover etc.)
        var minDate = DateTime.Today.AddDays(-MaxNavigationDays);
        var maxDate = DateTime.Today.AddDays(MaxNavigationDays);
        _prevButton.Enabled = _selectedDate > minDate;
        _nextButton.Enabled = _selectedDate < maxDate;
        _prevButton.ForeColor = _prevButton.Enabled
            ? Color.FromArgb(140, 145, 160)
            : Color.FromArgb(60, 64, 74);
        _nextButton.ForeColor = _nextButton.Enabled
            ? Color.FromArgb(140, 145, 160)
            : Color.FromArgb(60, 64, 74);

        _listPanel.ResumeLayout();
        ResumeLayout();

        UpdateCountdown();
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
            var text = _nowSepLabelText;
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
        string durationText;
        string timeText;
        if (meeting.IsAllDay)
        {
            // "Day" for a single-day event, "Day +N" for multi-day (N = remaining days after selected date)
            int remainingDays = (meeting.End.Date - _selectedDate.Date).Days - 1;
            durationText = remainingDays > 0 ? $"{remainingDays + 1} days" : "all day";
            timeText = remainingDays > 0 ? $"Day +{remainingDays}" : "Day";
        }
        else
        {
            var duration = meeting.End - meeting.Start;
            bool isMultiDay = meeting.End.Date > meeting.Start.Date;

            if (isMultiDay)
            {
                bool isStartDay = meeting.Start.Date == _selectedDate.Date;
                bool isEndDay   = meeting.End.Date   == _selectedDate.Date;

                if (isStartDay)
                {
                    // "10:00 +1d" — start time + how many days it continues
                    int dayDiff = (meeting.End.Date - meeting.Start.Date).Days;
                    timeText = $"{meeting.Start:HH:mm} +{dayDiff}d";
                    durationText = duration.TotalHours >= 1
                        ? (duration.Minutes > 0 ? $"{(int)duration.TotalHours}h {duration.Minutes}m" : $"{(int)duration.TotalHours}h")
                        : $"{(int)duration.TotalMinutes} min";
                }
                else if (isEndDay)
                {
                    // "→ 09:00" — arrow + end time on the last day
                    timeText = $"→ {meeting.End:HH:mm}";
                    durationText = string.Empty;
                }
                else
                {
                    // Middle day — same style as all-day
                    int remainingDays = (meeting.End.Date - _selectedDate.Date).Days - 1;
                    timeText = remainingDays > 0 ? $"Day +{remainingDays}" : "Day";
                    durationText = "all day";
                }
            }
            else
            {
                durationText = duration.TotalHours >= 1
                    ? (duration.Minutes > 0 ? $"{(int)duration.TotalHours}h {duration.Minutes}m" : $"{(int)duration.TotalHours}h")
                    : $"{(int)duration.TotalMinutes} min";
                timeText = $"{meeting.Start:HH:mm}–{meeting.End:HH:mm}";
            }
        }

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
            Font = new Font("Segoe UI", 8, meeting.IsAllDay ? FontStyle.Italic : FontStyle.Regular),
            ForeColor = Color.FromArgb(180, 180, 195),
            Text = timeText,
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

        // All rows are clickable to open the detail flyout (regardless of icon buttons)
        var clickHandler = (EventHandler)((_, _) => HandleRowClick(meeting));
        row.Click          += clickHandler;
        timeLabel.Click    += clickHandler;
        subjectLabel.Click += clickHandler;
        accountLabel.Click += clickHandler;
        durationLabel.Click += clickHandler;
        locationLabel.Click += clickHandler;
        row.Cursor = Cursors.Hand;

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
            acceptBtn.Click += async (_, _) =>
            {
                try { await _reminderService.RespondToMeetingAsync(meeting.Id, meeting.Start, true); }
                catch { }
                RefreshAgenda();
            };
            row.Controls.Add(acceptBtn);

            var declineBtn = MaterialIcons.MakeButton(MaterialIcons.ThumbDown, declineLeft, iconTop, iconSize,
                Color.FromArgb(210, 80, 90), rowBg);
            declineBtn.Click += async (_, _) =>
            {
                try { await _reminderService.RespondToMeetingAsync(meeting.Id, meeting.Start, false); }
                catch { }
                RefreshAgenda();
            };
            row.Controls.Add(declineBtn);
        }

        return row;
    }

    private async void HandleRowClick(ReminderMeeting meeting)
    {
        // Toggle: clicking the same row again closes the flyout
        if (_activeDetailForm is { IsDisposed: false } f && f.Tag is string id && id == meeting.Id)
        {
            _activeDetailForm.Close();
            _activeDetailForm = null;
            return;
        }

        _activeDetailForm?.Close();
        _activeDetailForm = null;

        var details = await _reminderService.GetMeetingDetailsAsync(meeting.Id, meeting.Start)
            ?? new MeetingDetails(meeting.DisplaySubject, meeting.Start, meeting.End,
                                  string.Empty, meeting.Location, meeting.Body, string.Empty, []);

        var anchorPoint = new Point(Right, Top);
        var flyout = new MeetingDetailForm(details, anchorPoint, async () => await _reminderService.OpenInOutlookAsync(meeting.Id, meeting.Start));
        flyout.Tag = meeting.Id;
        flyout.FormClosed += (_, _) => { if (_activeDetailForm == flyout) _activeDetailForm = null; };

        _activeDetailForm = flyout;
        flyout.Show(this); // owned by AgendaForm — won't trigger AgendaForm.Deactivate
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

    private static string GetDayRelativeLabel(DateTime selectedDate)
    {
        int daysFromToday = (selectedDate.Date - DateTime.Today).Days;
        return daysFromToday switch
        {
            0  => "NOW",
            1  => "tomorrow",
            -1 => "1 day ago",
            > 1 => $"in {daysFromToday} days",
            < -1 => $"{-daysFromToday} days ago"
        };
    }

    private void UpdateCountdown()
    {
        var now = DateTime.Now;
        int daysFromToday = (_selectedDate.Date - DateTime.Today).Days;

        // Update the separator label text (left side)
        var newLabelText = GetDayRelativeLabel(_selectedDate);
        bool labelChanged = _nowSepLabelText != newLabelText;
        _nowSepLabelText = newLabelText;

        string text;
        Color color;

        if (daysFromToday != 0)
        {
            // Non-today view — no live countdown
            text = string.Empty;
            color = Color.FromArgb(80, 85, 100);
        }
        else
        {
            var next = _cache.All
                .Where(m => !m.IsCancelled && !m.IsOutlookSynced && m.Start.Date == now.Date && m.Start > now)
                .OrderBy(m => m.Start)
                .FirstOrDefault();

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
        }

        bool colorChanged = _nowSepColor != color;
        _nowSepColor = color;
        if (_nowSepCountdownLabel is not null)
        {
            _nowSepCountdownLabel.Text = text;
            _nowSepCountdownLabel.ForeColor = color;
        }
        if (colorChanged || labelChanged)
            _nowSepPanel?.Invalidate();
    }

    private void ChangeSelectedDate(int dayOffset)
    {
        var newDate = _selectedDate.AddDays(dayOffset);
        var minDate = DateTime.Today.AddDays(-MaxNavigationDays);
        var maxDate = DateTime.Today.AddDays(MaxNavigationDays);

        // Clamp to ±7 days — never navigate outside this range
        if (newDate < minDate || newDate > maxDate)
            return;

        ApplySelectedDate(newDate);
    }

    private void ApplySelectedDate(DateTime newDate)
    {
        var minDate = DateTime.Today.AddDays(-MaxNavigationDays);
        var maxDate = DateTime.Today.AddDays(MaxNavigationDays);

        _selectedDate = newDate;
        _dateLabel.Text = _selectedDate.ToString("dddd, d MMMM");

        // Update day offset indicator
        int daysFromToday = (_selectedDate.Date - DateTime.Today).Days;
        if (daysFromToday > 0)
        {
            _dayOffsetLabel.Text = $"+{daysFromToday}";
            _dayOffsetLabel.ForeColor = Color.FromArgb(100, 160, 230);
            _dayOffsetLabel.Visible = true;
        }
        else if (daysFromToday < 0)
        {
            _dayOffsetLabel.Text = daysFromToday.ToString();
            _dayOffsetLabel.ForeColor = Color.FromArgb(140, 145, 160);
            _dayOffsetLabel.Visible = true;
        }
        else
        {
            _dayOffsetLabel.Visible = false;
        }

        // Enable/disable nav buttons at the limits
        _prevButton.Enabled = _selectedDate > minDate;
        _nextButton.Enabled = _selectedDate < maxDate;
        _prevButton.ForeColor = _prevButton.Enabled
            ? Color.FromArgb(140, 145, 160)
            : Color.FromArgb(60, 64, 74);
        _nextButton.ForeColor = _nextButton.Enabled
            ? Color.FromArgb(140, 145, 160)
            : Color.FromArgb(60, 64, 74);

        _lastMeetingsFingerprint = string.Empty; // Force refresh
        RefreshAgenda();
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        base.OnFormClosed(e);
        _fadeTimer?.Stop();
        _fadeTimer?.Dispose();
        _countdownTimer.Dispose();
        _ownerHandle.Dispose();
    }
}
