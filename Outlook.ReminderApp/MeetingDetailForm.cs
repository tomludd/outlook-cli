using System.Runtime.Versioning;

namespace Outlook.ReminderApp;

[SupportedOSPlatform("windows")]
internal sealed class MeetingDetailForm : Form
{
    private const int WsExAppWindow  = 0x00040000;
    private const int WsExToolWindow = 0x00000080;

    private const int FormWidth   = 760;

    protected override CreateParams CreateParams
    {
        get
        {
            var cp = base.CreateParams;
            cp.ExStyle |= WsExToolWindow;
            cp.ExStyle &= ~WsExAppWindow;
            return cp;
        }
    }

    public MeetingDetailForm(MeetingDetails details, Point anchorPoint, Action openInOutlook)
    {
        AutoScaleMode    = AutoScaleMode.None;
        FormBorderStyle  = FormBorderStyle.None;
        StartPosition    = FormStartPosition.Manual;
        ShowInTaskbar    = false;
        BackColor        = Color.FromArgb(30, 34, 44);
        ForeColor        = Color.WhiteSmoke;
        Width            = FormWidth;

        var workArea = Screen.FromPoint(anchorPoint).WorkingArea;
        int maxFormHeight = workArea.Height * 70 / 100;
        BuildContent(details, openInOutlook, maxFormHeight);
        PositionFlyout(anchorPoint);

        Deactivate += (_, _) =>
        {
            // Defer so the owner's row-click can fire and call HandleRowClick
            // before we decide whether to close.
            BeginInvoke(() =>
            {
                if (IsDisposed) return;
                // Stay open if this flyout is still the active form
                if (Form.ActiveForm == this) return;
                // Stay open if the owner (AgendaForm) got focus back — user is switching rows
                if (Owner is Form owner && owner == Form.ActiveForm) return;
                Close();
            });
        };
    }

    private void BuildContent(MeetingDetails details, Action openInOutlook, int maxFormHeight)
    {
        // Layout (FormWidth = 760):
        // [4px accent][16px pad][body 480px ≈80 chars][8px][1px div][8px][attendees 231px][12px]
        const int colLeft      = 20;    // accent(4) + pad(16)
        const int bodyColWidth = 480;   // ≈80 chars Segoe UI 8.5pt
        const int vertDivX     = colLeft + bodyColWidth + 8;   // 508
        const int attColLeft   = vertDivX + 1 + 8;             // 517
        const int attColWidth  = FormWidth - attColLeft - 12;  // 231
        const int fullWidth    = FormWidth - colLeft - 12;     // 728

        var accentColor  = Color.FromArgb(80, 140, 200);
        var dividerColor = Color.FromArgb(50, 55, 70);
        var dimText      = Color.FromArgb(140, 145, 160);
        var bodyText     = Color.FromArgb(200, 200, 210);

        SuspendLayout();

        var scrollPanel = new Panel
        {
            Left = 0, Top = 0, Width = FormWidth,
            AutoScroll = true,
            BackColor = Color.FromArgb(30, 34, 44),
        };
        Controls.Add(scrollPanel);
        scrollPanel.SuspendLayout();

        var accent = new Panel { Left = 0, Top = 0, Width = 4, BackColor = accentColor };
        scrollPanel.Controls.Add(accent);

        int y = 12;

        // ── Subject + Open in Outlook button (top-right) ──
        const int openBtnWidth = 130;
        var openBtn = new Button
        {
            AutoSize = false,
            Left = FormWidth - openBtnWidth - 12, Top = y, Width = openBtnWidth, Height = 22,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 8f, FontStyle.Regular),
            ForeColor = Color.FromArgb(80, 140, 200),
            BackColor = Color.FromArgb(30, 34, 44),
            Text = "Open in Outlook →",
            TextAlign = ContentAlignment.MiddleRight,
            Cursor = Cursors.Hand, TabStop = false
        };
        openBtn.FlatAppearance.BorderSize = 0;
        openBtn.FlatAppearance.MouseOverBackColor = Color.FromArgb(38, 43, 55);
        openBtn.FlatAppearance.MouseDownBackColor = Color.FromArgb(45, 50, 65);
        openBtn.Click += (_, _) => openInOutlook();
        scrollPanel.Controls.Add(openBtn);

        scrollPanel.Controls.Add(new Label
        {
            AutoSize = false, AutoEllipsis = true,
            Left = colLeft, Top = y, Width = fullWidth - openBtnWidth - 8, Height = 22,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = Color.WhiteSmoke, Text = details.Subject,
            TextAlign = ContentAlignment.MiddleLeft
        });
        y += 22 + 2;

        // ── Time range ──
        if (details.Start != DateTime.MinValue)
        {
            var dur = details.End - details.Start;
            string durText = dur.TotalHours >= 1
                ? (dur.Minutes > 0 ? $"{(int)dur.TotalHours}h {dur.Minutes}m" : $"{(int)dur.TotalHours}h")
                : $"{(int)dur.TotalMinutes} min";
            scrollPanel.Controls.Add(new Label
            {
                AutoSize = false,
                Left = colLeft, Top = y, Width = fullWidth, Height = 18,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Regular),
                ForeColor = dimText,
                Text = $"{details.Start:HH:mm} – {details.End:HH:mm}  ·  {durText}",
                TextAlign = ContentAlignment.MiddleLeft
            });
            y += 18;
        }

        y += 6;
        scrollPanel.Controls.Add(new Panel { Left = colLeft, Top = y, Width = fullWidth, Height = 1, BackColor = dividerColor });
        y += 9;

        // ── Organizer + Location (full width, above body columns) ──
        bool hasTopInfo = false;

        if (!string.IsNullOrWhiteSpace(details.Organizer))
        {
            scrollPanel.Controls.Add(new Label
            {
                AutoSize = false, AutoEllipsis = true,
                Left = colLeft, Top = y, Width = fullWidth, Height = 18,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Regular),
                ForeColor = bodyText, Text = $"Organizer: {details.Organizer}",
                TextAlign = ContentAlignment.MiddleLeft
            });
            y += 18 + 2;
            hasTopInfo = true;
        }

        var locationParts = details.Location
            .Split(['\r', '\n', ';', '|'], StringSplitOptions.RemoveEmptyEntries)
            .Select(p => p.Trim())
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .Where(p => !string.Equals(p, "Microsoft Teams Meeting", StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (locationParts.Count > 0)
        {
            scrollPanel.Controls.Add(new Label
            {
                AutoSize = false, AutoEllipsis = true,
                Left = colLeft, Top = y, Width = fullWidth, Height = 18,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Regular),
                ForeColor = bodyText, Text = $"Location: {string.Join(" | ", locationParts)}",
                TextAlign = ContentAlignment.MiddleLeft
            });
            y += 18 + 2;
            hasTopInfo = true;
        }

        if (hasTopInfo)
        {
            y += 4;
            scrollPanel.Controls.Add(new Panel { Left = colLeft, Top = y, Width = fullWidth, Height = 1, BackColor = dividerColor });
            y += 9;
        }

        int colStartY = y;
        int leftY     = y;
        int rightY    = y;

        // ── Left column: Body (HTML if available, plain text fallback) ──
        var trimmedHtml  = details.HtmlBody.Trim();
        var trimmedPlain = details.Body.Trim();
        bool hasBody = !string.IsNullOrEmpty(trimmedHtml) || !string.IsNullOrEmpty(trimmedPlain);

        if (hasBody)
        {
            const string darkCss =
                "<style>" +
                "body,html{background:#1e222c!important;color:#c8c8d2!important;" +
                "margin:0;padding:4px 6px;font-family:'Segoe UI',sans-serif;font-size:9pt;" +
                "overflow-x:hidden;word-wrap:break-word;overflow-wrap:break-word}" +
                "h1,h2,h3,h4,h5,h6{color:#e8e8f0!important;margin-top:0.4em;margin-bottom:0.2em}" +
                "h1{font-size:1.5em}h2{font-size:1.3em}h3{font-size:1.1em}" +
                "a,a:link,a:visited{color:#6ba3e8!important;text-decoration:underline;cursor:pointer}" +
                "table,td,th,div,p,span,blockquote{background:transparent!important}" +
                "</style>";

            string pageHtml;
            if (!string.IsNullOrEmpty(trimmedHtml))
            {
                // Strip the Teams boilerplate footer that Outlook injects:
                // "__________________ Join Microsoft Teams Meeting ..." block
                var cleaned = System.Text.RegularExpressions.Regex.Replace(
                    trimmedHtml,
                    @"_{5,}.*",
                    "",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase |
                    System.Text.RegularExpressions.RegexOptions.Singleline);

                // Unwrap Outlook SafeLinks: strip the tracking wrapper, keep display text
                cleaned = System.Text.RegularExpressions.Regex.Replace(
                    cleaned,
                    @"<a\s[^>]*href=""https://[^""]*safelinks\.protection\.outlook\.com[^""]*""[^>]*>(.*?)</a>",
                    "$1",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase |
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                // Unwrap tel:/mailto: links — keep display text, drop the link
                cleaned = System.Text.RegularExpressions.Regex.Replace(
                    cleaned,
                    @"<a\s[^>]*href=""(?:tel|mailto):[^""]*""[^>]*>(.*?)</a>",
                    "$1",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase |
                    System.Text.RegularExpressions.RegexOptions.Singleline);

                // Collapse runs of whitespace-only content in the HTML body:
                // 2+ consecutive <br> → 1
                cleaned = System.Text.RegularExpressions.Regex.Replace(
                    cleaned,
                    @"(<br\s*/?>(\s|&nbsp;)*){2,}",
                    "<br>",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // 2+ consecutive empty <p> → removed
                cleaned = System.Text.RegularExpressions.Regex.Replace(
                    cleaned,
                    @"(<p[^>]*>(\s|&nbsp;)*</p>\s*){2,}",
                    "",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // 2+ consecutive empty <div> → removed
                cleaned = System.Text.RegularExpressions.Regex.Replace(
                    cleaned,
                    @"(<div[^>]*>(\s|&nbsp;)*</div>\s*){2,}",
                    "",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Join numbered list items split by <br>: "1.<br>text" → "1. text"
                cleaned = System.Text.RegularExpressions.Regex.Replace(
                    cleaned,
                    @"(\d+\.)\s*(<br\s*/?>\s*)+([^<\r\n])",
                    "$1 $3",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // Inject dark CSS into existing <head> tag
                int headStart = cleaned.IndexOf("<head", StringComparison.OrdinalIgnoreCase);
                if (headStart >= 0)
                {
                    int headClose = cleaned.IndexOf('>', headStart);
                    pageHtml = headClose >= 0
                        ? cleaned.Insert(headClose + 1, darkCss)
                        : cleaned;
                }
                else
                {
                    pageHtml = $"<html><head>{darkCss}</head><body>{cleaned}</body></html>";
                }
            }
            else
            {
                // Strip Teams boilerplate footer (starts with a run of underscores)
                var collapsedPlain = System.Text.RegularExpressions.Regex.Replace(
                    trimmedPlain,
                    @"_{5,}.*",
                    "",
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                // Strip SafeLinks URLs: <https://...safelinks...> and bare URLs
                collapsedPlain = System.Text.RegularExpressions.Regex.Replace(
                    collapsedPlain,
                    @"\s*<https://[^>]*safelinks\.protection\.outlook\.com[^>]*>",
                    "",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Strip <tel:...> and <mailto:...> angle-bracket references
                collapsedPlain = System.Text.RegularExpressions.Regex.Replace(
                    collapsedPlain,
                    @"\s*<(?:tel|mailto):[^>]*>",
                    "",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Collapse 2+ consecutive blank lines down to 1
                collapsedPlain = System.Text.RegularExpressions.Regex.Replace(
                    collapsedPlain,
                    @"(\r?\n){3,}",
                    "\n");
                // Join numbered list items split across lines: "1.\ntext" or "1.\n\ntext" → "1. text"
                collapsedPlain = System.Text.RegularExpressions.Regex.Replace(
                    collapsedPlain,
                    @"(\d+\.)\s*\r?\n\s*([^\r\n])",
                    "$1 $2");
                collapsedPlain = collapsedPlain.TrimEnd();
                var escaped = collapsedPlain
                    .Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
                    .Replace("\r\n", "<br>").Replace("\n", "<br>");
                pageHtml = $"<html><head>{darkCss}</head><body>{escaped}</body></html>";
            }

            var browser = new WebBrowser
            {
                Left = colLeft, Top = y, Width = bodyColWidth, Height = 360,
                ScrollBarsEnabled = false,
                IsWebBrowserContextMenuEnabled = false,
                WebBrowserShortcutsEnabled = false,
                TabStop = false
            };
            // Open links in the default browser instead of navigating inside the control
            browser.Navigating += (_, e) =>
            {
                if (e.Url is null || e.Url.Scheme == "about" || e.Url.Scheme == "res") return;
                e.Cancel = true;
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = e.Url.ToString(),
                    UseShellExecute = true
                });
            };
            // After the document loads, resize the browser to its actual content height and
            // shift all controls below it by the same delta so the layout stays correct.
            browser.DocumentCompleted += (_, _) =>
            {
                if (browser.IsDisposed || browser.Document?.Body is null) return;
                int docHeight = browser.Document.Body.ScrollRectangle.Height + 8;
                int delta = docHeight - browser.Height;
                if (delta == 0) return;

                browser.Height = docHeight;

                // Shift only left-column and full-width controls below the browser's original bottom.
                // Attendee controls (right column) are independent and must NOT be shifted.
                int originalBottom = browser.Bottom - delta; // bottom before resize
                foreach (Control c in scrollPanel.Controls)
                {
                    if (c != browser && c.Top > originalBottom && c.Left < attColLeft)
                        c.Top += delta;
                }

                // Extend the vertical accent bar to match new total height
                int newTotal = scrollPanel.Controls.Cast<Control>().Max(c => c.Bottom) + 10;
                foreach (Control c in scrollPanel.Controls)
                {
                    if (c is Panel { Width: 4, Left: 0 }) // accent bar
                        c.Height = newTotal;
                }
                scrollPanel.Height = Math.Min(newTotal, maxFormHeight);
                Height = scrollPanel.Height;
            };
            browser.DocumentText = pageHtml;
            scrollPanel.Controls.Add(browser);
            leftY = browser.Bottom + 4;
        }

        // ── Right column: Attendees ──
        var groups = new (string Label, Color Color, string Key)[]
        {
            ("✓  Accepted",    Color.FromArgb(80, 190, 120),  "Accepted"),
            ("?  Tentative",   Color.FromArgb(255, 195, 60),  "Tentative"),
            ("✗  Declined",    Color.FromArgb(220, 80, 90),   "Declined"),
            ("–  No response", Color.FromArgb(120, 125, 140), "Not Responded"),
        };

        // Build a set of names that appear more than once across all attendees so we can
        // disambiguate them with their email domain.
        var duplicateNames = details.Attendees
            .GroupBy(a => a.Name, StringComparer.OrdinalIgnoreCase)
            .Where(g => g.Count() > 1)
            .Select(g => g.Key)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        static string DisplayName(AttendeeInfo a, HashSet<string> dupes)
        {
            if (!dupes.Contains(a.Name)) return a.Name;
            // Extract domain from email: "tommy@epinova.no" → "epinova"
            var email = a.Email ?? string.Empty;
            var atIdx = email.IndexOf('@');
            if (atIdx < 0) return a.Name;
            var domain = email[(atIdx + 1)..];
            var dotIdx = domain.IndexOf('.');
            var domainLabel = dotIdx > 0 ? domain[..dotIdx] : domain;
            return string.IsNullOrEmpty(domainLabel) ? a.Name : $"{a.Name} ({domainLabel})";
        }

        bool hasAnyAttendees = details.Attendees.Count > 0;
        foreach (var (groupLabel, groupColor, statusKey) in groups)
        {
            var members = details.Attendees
                .Where(a => string.Equals(a.ResponseStatus, statusKey, StringComparison.OrdinalIgnoreCase))
                .OrderBy(a => a.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (members.Count == 0) continue;

            scrollPanel.Controls.Add(new Label
            {
                AutoSize = false,
                Left = attColLeft, Top = rightY, Width = attColWidth, Height = 18,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                ForeColor = groupColor,
                Text = $"{groupLabel}  ({members.Count})",
                TextAlign = ContentAlignment.MiddleLeft
            });
            rightY += 18 + 2;

            foreach (var attendee in members)
            {
                scrollPanel.Controls.Add(new Label
                {
                    AutoSize = false, AutoEllipsis = true,
                    Left = attColLeft + 8, Top = rightY, Width = attColWidth - 8, Height = 17,
                    Font = new Font("Segoe UI", 8.5f, FontStyle.Regular),
                    ForeColor = bodyText, Text = DisplayName(attendee, duplicateNames),
                    TextAlign = ContentAlignment.MiddleLeft
                });
                rightY += 17 + 1;
            }
            rightY += 5;
        }

        if (!hasAnyAttendees)
        {
            scrollPanel.Controls.Add(new Label
            {
                AutoSize = false,
                Left = attColLeft, Top = rightY, Width = attColWidth, Height = 18,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Regular),
                ForeColor = dimText, Text = "No attendee data",
                TextAlign = ContentAlignment.MiddleLeft
            });
            rightY += 18;
        }

        // ── Vertical separator ──
        int colEndY = Math.Max(leftY, rightY) + 4;
        scrollPanel.Controls.Add(new Panel
        {
            Left = vertDivX, Top = colStartY, Width = 1, Height = colEndY - colStartY,
            BackColor = dividerColor
        });

        y = colEndY + 4;

        accent.Height = y;
        scrollPanel.Height = Math.Min(y, maxFormHeight);
        Height = scrollPanel.Height;

        scrollPanel.ResumeLayout(false);
        ResumeLayout(false);
    }

    private void PositionFlyout(Point anchorPoint)
    {
        var screen = Screen.FromPoint(anchorPoint).WorkingArea;

        int x = anchorPoint.X + 6;
        int y = anchorPoint.Y;

        if (x + Width > screen.Right)
            x = anchorPoint.X - Width - 6;

        if (y + Height > screen.Bottom)
            y = screen.Bottom - Height;

        if (y < screen.Top)
            y = screen.Top;

        Left = x;
        Top  = y;
    }
}
