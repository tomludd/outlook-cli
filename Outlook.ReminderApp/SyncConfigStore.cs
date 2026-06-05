using System.Text.Json;

namespace Outlook.ReminderApp;

internal sealed class SyncConfigStore
{
    private const string FileName = "sync-rules.json";

    public string ConfigPath
    {
        get
        {
            var root = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var dir = Path.Combine(root, "Outlook.ReminderApp");
            return Path.Combine(dir, FileName);
        }
    }

    public List<SyncRule> Load()
    {
        var path = ConfigPath;
        if (!File.Exists(path))
        {
            return new List<SyncRule>();
        }

        try
        {
            var json = File.ReadAllText(path);
            var rules = JsonSerializer.Deserialize<List<SyncRule>>(json);
            return rules ?? new List<SyncRule>();
        }
        catch
        {
            return new List<SyncRule>();
        }
    }

    public void Save(IEnumerable<SyncRule> rules)
    {
        var path = ConfigPath;
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
        {
            Directory.CreateDirectory(dir);
        }

        var json = JsonSerializer.Serialize(rules, new JsonSerializerOptions
        {
            WriteIndented = true
        });
        File.WriteAllText(path, json);
    }
}
