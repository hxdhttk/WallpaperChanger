using System.Text.Json;
using System.Text.Json.Serialization;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Jpeg;
using SixLabors.ImageSharp.Processing;
using Windows.Win32;
using Windows.Win32.UI.Shell;

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
    var wallpaper = (IDesktopWallpaper)new DesktopWallpaper();

    wallpaper.SetPosition(DESKTOP_WALLPAPER_POSITION.DWPOS_FILL);
    wallpaper.SetWallpaper(null, imagePath);
}

static async Task StartChangeWallpaperAsync(string imageFolder, int intervalInSeconds)
{
    var intervalSpan = TimeSpan.FromSeconds(intervalInSeconds);

    while (true)
    {
        var images = Directory.GetFiles(imageFolder);
        images = [.. images.OrderBy(_ => Random.Shared.Next())];

        foreach (var image in images)
        {
            if (File.Exists(image))
            {
                try
                {
                    var transcodedImage = await TranscodeAsync(image);
                    ChangeWallpaper(transcodedImage);
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

static async Task<string> TranscodeAsync(string image)
{
    var fileSizeInBytes = new FileInfo(image).Length;
    if (fileSizeInBytes <= Constants.MaxFileSizeInBytes)
    {
        return image;
    }

    var appPathRoot = AppContext.BaseDirectory;
    var transcodedImage = Path.Combine(appPathRoot, Path.GetFileName(image));
    if (File.Exists(transcodedImage))
    {
        return transcodedImage;
    }

    var imageObj = await Image.LoadAsync(image);

    var imageMaxDim = Math.Max(imageObj.Width, imageObj.Height);
    if (imageMaxDim > Constants.MaxDim)
    {
        var scale = Constants.MaxDim / (float)imageMaxDim;
        var newWidth = (int)(imageObj.Width * scale);
        var newHeight = (int)(imageObj.Height * scale);
        imageObj.Mutate(x => x.Resize(newWidth, newHeight));
    }

    await imageObj.SaveAsJpegAsync(
        transcodedImage,
        new JpegEncoder { Quality = Constants.TranscodeQuality }
    );

    return transcodedImage;
}

public class Config
{
    public string? ImageFolder { get; set; }

    public int IntervalInSeconds { get; set; } = 30;
}

public static class Constants
{
    public const int MaxDim = 4096;

    public const int MaxFileSizeInBytes = 15 * 1024 * 1024;

    public const int TranscodeQuality = 85;
}

[JsonSerializable(typeof(Config))]
internal partial class SourceGenerationContext : JsonSerializerContext { }
