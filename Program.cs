using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Win32;
using Windows.Win32;
using Windows.Win32.Foundation;
using Windows.Win32.Graphics.Gdi;
using Windows.Win32.UI.WindowsAndMessaging;

var configFileContent = File.ReadAllBytes("Config.json");
var config = JsonSerializer.Deserialize(configFileContent, SourceGenerationContext.Default.Config);

ArgumentNullException.ThrowIfNull(config);
ArgumentException.ThrowIfNullOrEmpty(config.ImageFolder);

await StartChangeWallpaperAsync(
    Environment.ExpandEnvironmentVariables(config.ImageFolder),
    config.IntervalInSeconds
);

static unsafe void ChangeWallpaper(string imagePath)
{
    // 设置桌面背景的样式
    RegistryKey? key = Registry.CurrentUser.OpenSubKey(@"Control Panel\Desktop", true);
    ArgumentNullException.ThrowIfNull(key);

    key.SetValue(@"WallpaperStyle", "10");
    key.SetValue(@"TileWallpaper", "0");

    // 设置桌面背景
    var imagePathPtr = Marshal.StringToHGlobalAuto(imagePath);
    PInvoke.SystemParametersInfo(
        SYSTEM_PARAMETERS_INFO_ACTION.SPI_SETDESKWALLPAPER,
        0,
        imagePathPtr.ToPointer(),
        SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS.SPIF_UPDATEINIFILE
            | SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS.SPIF_SENDWININICHANGE
    );
    Marshal.FreeHGlobal(imagePathPtr);

    // 刷新桌面
    var desktop = PInvoke.GetDesktopWindow();
    PInvoke.RedrawWindow(
        desktop,
        (RECT*)0,
        (HRGN)IntPtr.Zero,
        REDRAW_WINDOW_FLAGS.RDW_INVALIDATE
            | REDRAW_WINDOW_FLAGS.RDW_UPDATENOW
            | REDRAW_WINDOW_FLAGS.RDW_ERASE
            | REDRAW_WINDOW_FLAGS.RDW_ALLCHILDREN
    );
}

static async Task StartChangeWallpaperAsync(string imageFolder, int intervalInSeconds)
{
    var intervalSpan = TimeSpan.FromSeconds(intervalInSeconds);

    while (true)
    {
        var images = Directory.GetFiles(imageFolder);
        images =  [.. images.OrderBy(_ => Random.Shared.Next())];

        foreach (var image in images)
        {
            if (File.Exists(image))
            {
                try
                {
                    ChangeWallpaper(image);
                    await Task.Delay(intervalSpan);
                }
                catch
                {
                    // Ignore exceptions
                }
            }
        }
    }
}

public class Config
{
    public string? ImageFolder { get; set; }

    public int IntervalInSeconds { get; set; } = 30;
}

[JsonSerializable(typeof(Config))]
internal partial class SourceGenerationContext : JsonSerializerContext { }
