using System.Drawing.Text;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace Outlook.ReminderApp;

[SupportedOSPlatform("windows")]
internal static class MaterialIcons
{
    public const string VideoCall   = "\ue070"; // join meeting
    public const string ChatBubble  = "\ue0ca"; // open chat
    public const string ThumbUp     = "\ue8dc"; // accept
    public const string ThumbDown   = "\ue8db"; // decline
    public const string Close       = "\ue5cd"; // dismiss

    private static readonly PrivateFontCollection _fontCollection = Load();
    public  static readonly FontFamily Family = _fontCollection.Families[0];

    private static PrivateFontCollection Load()
    {
        var pfc = new PrivateFontCollection();
        var asm = typeof(MaterialIcons).Assembly;
        using var stream = asm.GetManifestResourceStream(
            "Outlook.ReminderApp.Resources.MaterialIcons-Regular.ttf")!;
        var data = new byte[stream.Length];
        _ = stream.Read(data, 0, data.Length);
        var handle = GCHandle.Alloc(data, GCHandleType.Pinned);
        try { pfc.AddMemoryFont(handle.AddrOfPinnedObject(), data.Length); }
        finally { handle.Free(); }
        return pfc;
    }

    public static Label MakeButton(string glyph, int x, int y, int size, Color color, Color bgColor, float fontScale = 0.52f)
    {
        return new Label
        {
            Left = x, Top = y, Width = size, Height = size,
            Text = glyph,
            Font = new Font(Family, size * fontScale, FontStyle.Regular, GraphicsUnit.Pixel),
            TextAlign = ContentAlignment.MiddleCenter,
            Cursor = Cursors.Hand,
            BackColor = bgColor,
            ForeColor = color,
            UseCompatibleTextRendering = true
        };
    }
}
