using System.Text.Json;

public class SettingsService
{
    private readonly string _path;
    private readonly object _lock = new();

    public SettingsService()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "SaveAsPDF");
        Directory.CreateDirectory(dir);
        _path = Path.Combine(dir, "settings.json");
    }

    public AppSettings Load()
    {
        lock (_lock)
        {
            if (!File.Exists(_path)) return new AppSettings();
            try
            {
                var json = File.ReadAllText(_path);
                return JsonSerializer.Deserialize<AppSettings>(
                    json,
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
                       ?? new AppSettings();
            }
            catch { return new AppSettings(); }
        }
    }

    public void Save(AppSettings settings)
    {
        lock (_lock)
        {
            File.WriteAllText(_path, JsonSerializer.Serialize(
                settings,
                new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                }));
        }
    }
}
